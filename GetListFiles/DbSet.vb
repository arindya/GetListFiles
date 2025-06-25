Imports Oracle.ManagedDataAccess.Client

Public Class DbSet

    Private Sub DbSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnOn_Click(sender As Object, e As EventArgs) Handles btnOn.Click
        Dim userId As String = My.Settings.OracleUserId
        Dim passwordx As String = My.Settings.OraclePassword
        Dim host As String = My.Settings.OracleHost
        Dim port As String = My.Settings.OraclePort
        Dim serviceName As String = My.Settings.OracleService
        If String.IsNullOrEmpty(txUid.Text) = False Or String.IsNullOrEmpty(txPwd.Text) = False Or String.IsNullOrEmpty(txSvr.Text) = False Then
            userID = txUid.Text
            password = txPwd.Text
            dSource = txSvr.Text

            Using conn As New OracleConnection
                conn.ConnectionString = String.Format(
    "User Id={0};Password={1};Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={2})(PORT={3}))(CONNECT_DATA=(SERVICE_NAME={4})))",
    userId, passwordx, host, port, serviceName)

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