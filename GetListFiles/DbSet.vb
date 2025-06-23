Imports Oracle.ManagedDataAccess.Client

Public Class DbSet

    Private Sub DbSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnOn_Click(sender As Object, e As EventArgs) Handles btnOn.Click
        If String.IsNullOrEmpty(txUid.Text) = False Or String.IsNullOrEmpty(txPwd.Text) = False Or String.IsNullOrEmpty(txSvr.Text) = False Then
            userID = txUid.Text
            password = txPwd.Text
            dSource = txSvr.Text

            Using conn As New OracleConnection
                conn.ConnectionString = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

                Try
                    conn.Open()
                    My.Settings.UserDB = txUid.Text
                    My.Settings.PassDB = txPwd.Text
                    My.Settings.DSource = txSvr.Text
                    My.Settings.Save()
                    MessageBox.Show("Database connection saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    flag = "success"
                    flag = "success"
                    Close()
                Catch ex As Exception
                    flag = "false"
                    MessageBox.Show("Invalid Datasource, please ensure that you input the correct datasource kondisi 2 salah", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                conn.Close()
            End Using
        Else
            MessageBox.Show("Please fill all the input field", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
End Class