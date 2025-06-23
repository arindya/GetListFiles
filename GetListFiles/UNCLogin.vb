Public Class UNCLogin
    Private Sub UNCLogin_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        tbDomain.Text = My.Settings.uncDomain
        tbUsername.Text = My.Settings.uncUsername
        tbPassword.Text = My.Settings.pwunc
        tbPath.Text = My.Settings.DigdatHost
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        Me.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        My.Settings.uncDomain = tbDomain.Text
        My.Settings.uncUsername = tbUsername.Text
        My.Settings.pwunc = tbPassword.Text
        My.Settings.DigdatHost = tbPath.Text
        My.Settings.Save()
        Application.DoEvents()
        Using unc As New UNCAccess()
            If Not String.IsNullOrEmpty(tbUsername.Text) And Not String.IsNullOrEmpty(tbPassword.Text) And Not String.IsNullOrEmpty(tbPath.Text) Then
                If unc.NetUseWithCredentials(tbPath.Text, tbUsername.Text, tbDomain.Text, tbPassword.Text) Then
                    MessageBox.Show("Successfully Connected!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Close()
                Else
                    Me.Cursor = Cursors.[Default]
                    Select Case unc.LastError
                        Case 1219
                            MessageBox.Show("Already logged in", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Me.Close()
                        Case 67
                            MessageBox.Show("The network name cannot be found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Case 58
                            MessageBox.Show("The specified server cannot perform the requested operation", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Case 1326
                            MessageBox.Show("The username or password is incorrect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Select
                    'MessageBox.Show("Failed to connect to " + tbPath.Text + vbCr & vbLf & "LastError = " + unc.LastError.ToString(), "Failed to connect", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End If
            Else
                MessageBox.Show("Please ensure that you fill all fields", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Using
        Me.Cursor = Cursors.[Default]
        Me.Enabled = True
    End Sub
End Class