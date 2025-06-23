Imports Oracle.ManagedDataAccess.Client
Imports System.Security.Cryptography
Imports System.Text

Public Class Login
    Private Sub Login_load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Settings.uname
        TextBox2.Text = My.Settings.pass
    End Sub
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        My.Settings.uname = TextBox1.Text
        My.Settings.pass = TextBox2.Text
        My.Settings.Save()
        Me.Enabled = False
        Dim username As String = TextBox1.Text
        Dim password As String = MD5Hash(TextBox2.Text)
        Dim conString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"
        'Dim conString As String = "Data Source=" + My.Settings.DSource + ";User ID=" + My.Settings.UserDB + ";Password=" + My.Settings.PassDB + ""
        Dim query As String = "select user_id, user_pass from users where user_id = :user_id and user_pass = :password"
        Dim ds As DataSet = New DataSet()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.Add(":user_id", OracleDbType.Varchar2).Value = username
                    .Parameters.Add(":password", OracleDbType.Varchar2).Value = password
                End With
                Try
                    conn.Open()
                    Dim da As OracleDataAdapter = New OracleDataAdapter(comm)
                    da.Fill(ds, "login")
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    Return
                End Try
            End Using
        End Using


        If String.IsNullOrEmpty(username) = True Or String.IsNullOrEmpty(password) = True Then
            MessageBox.Show("Username and Password cannot be empty", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf String.IsNullOrEmpty(username) = True Then
            MessageBox.Show("Username is empty", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        ElseIf String.IsNullOrEmpty(password) = True Then
            MessageBox.Show("Password is empty", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim num_rows As Integer = ds.Tables(0).Rows.Count
        If num_rows = 0 Then
            MessageBox.Show("The user is not recongized", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If ds.Tables(0).Rows(0)("user_pass").ToString <> password Then
            MessageBox.Show("Invalid username or password", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Else
            Me.DialogResult = DialogResult.OK
            Close()
        End If
        Me.Enabled = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Function MD5Hash(ByVal MD5Data As String)
        Dim AE As ASCIIEncoding = New ASCIIEncoding()
        Dim data() As Byte = AE.GetBytes(MD5Data)
        Dim md5 As MD5 = New MD5CryptoServiceProvider()
        Dim result() As Byte = md5.ComputeHash(data)
        Dim strResult As String = ""
        For x As Integer = 0 To result.Length - 1
            strResult = strResult + result(x).ToString
        Next
        Return strResult
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.Reset()
        Close()
    End Sub
End Class