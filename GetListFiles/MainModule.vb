Module MainModule
    Public flag As String
    Public userID As String
    Public password As String
    Public dSource As String

    Public Sub Main()
        'My.Settings.Reset()
        Application.EnableVisualStyles()
tes:
        If My.Settings.UserDB = String.Empty And My.Settings.PassDB = String.Empty And My.Settings.DSource = String.Empty Then
            Dim result As String = DbSet.ShowDialog()
            If result = DialogResult.OK Then
                If String.IsNullOrEmpty(userID) = False And String.IsNullOrEmpty(password) = False And String.IsNullOrEmpty(dSource) = False Then
                    If flag = "success" Then
                        Application.Run(GLV)
                    ElseIf flag = "failed"
                        GoTo tes
                    End If
                Else
                    GoTo tes
                End If
            End If
        Else
            Application.Run(GLV)
        End If
    End Sub
End Module