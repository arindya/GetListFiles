'+-----------------------------+
'|  Created By   : R.Armayndo  |
'|  Created Date : 20120504    |
'+-----------------------------+
Imports System.IO
Public Class FindReplace
    Dim files() As String
    Dim respon As String

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click

        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        ' Description that displays above the dialog box control.

        MyFolderBrowser.Description = "Select the Folder"
        MyFolderBrowser.ShowNewFolderButton = False
        ' Sets the root folder where the browsing starts from
        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txBrowse.Text = MyFolderBrowser.SelectedPath
        End If
    End Sub

    Private Sub btnProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess.Click
        If txFind.Text <> "" And Trim(txBrowse.Text) <> "" Then
            If Directory.Exists(txBrowse.Text) Then
                TextBox1.Text = ""
                Dim cari As String = Trim(txFind.Text)
                If cbFindMod.Checked = True Then
                    files = Directory.GetFiles(txBrowse.Text, "*" + txFind.Text + "*", SearchOption.AllDirectories)
                Else
                    files = Directory.GetFiles(txBrowse.Text, "*" + txFind.Text + "*", SearchOption.TopDirectoryOnly)
                End If

                Dim x As Integer
                For x = 0 To files.Length - 1
                    System.Threading.Thread.Sleep(100)
                    TextBox1.Text = Path.GetFileName(files(x).ToString) + vbCrLf + TextBox1.Text
                    TextBox1.Refresh()
                Next
                MsgBox("Searching complete, " + files.Length.ToString + " Files found", MsgBoxStyle.Information, "FINISH")
            Else
                MsgBox("Directory does not exist", MsgBoxStyle.Exclamation, "WARNING")
            End If

        Else
            MsgBox("Find text and Directory cannot be null", MsgBoxStyle.Exclamation, "WARNING")
        End If
    End Sub

    Private Sub FindReplace_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txBrowse.Text = Application.StartupPath.ToString
    End Sub

    Private Sub btnReplace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReplace.Click
        Dim namabaru, namalama, StrMsg, strPath As String
        Dim CountErr As Int16 = 0
        Dim CountSama As Int16 = 0
        Dim CountOK As Int16 = 0
        respon = MsgBox("There's " + (files.Length).ToString + " files found, " + vbCr + "Replace '" + txFind.Text.ToString + "' to '" + txReplace.Text.ToString + "'" + vbCr + "Do you wish to continued?", MsgBoxStyle.YesNo, "CONFIRMATION")
        Select Case respon
            Case vbYes
                Dim x As Integer

                ' Set up the progress bar's properties
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = files.Length
                ProgressBar1.Value = 0

                For x = 0 To files.Length - 1

                    ProgressBar1.Value = ProgressBar1.Value + 1
                    lbl_record.Text = x + 1 & " of " & files.Length.ToString & " Files"
                    lbl_record.Refresh()

                    System.Threading.Thread.Sleep(100)
                    strPath = Path.GetDirectoryName(files(x).ToString)
                    If cbCaseStv.Checked = True Then
                        namabaru = Path.GetFileName(files(x).ToString).Replace(txFind.Text, txReplace.Text)
                    Else
                        namabaru = Path.GetFileName(files(x).ToString.ToUpper).Replace(txFind.Text.ToUpper, txReplace.Text.ToUpper)
                    End If

                    namalama = Path.GetFileName(files(x).ToString)
                    If File.Exists(strPath + "\" + namabaru) Then
                        'do nothing
                        CountSama += 1
                    Else
                        Try
                            File.Move(strPath + "\" + namalama, strPath + "\" + namabaru)
                            CountOK += 1
                        Catch ex As Exception
                            CountErr += 1
                            MsgBox(ex.ToString, MsgBoxStyle.Critical, "ERROR")
                        End Try
                    End If
                Next
                StrMsg = "Finished, " + CountOK.ToString + " Files has been renamed."
                If CountSama <> 0 Then
                    StrMsg = StrMsg + vbCrLf + CountSama.ToString + " Files cannot be renamed because destination name been exist"
                End If
                If CountErr <> 0 Then
                    StrMsg = StrMsg + vbCrLf + CountErr.ToString + " Files cannot be renamed because unexpected ERROR"
                End If

                MsgBox(StrMsg, MsgBoxStyle.Information, "FINISH")
            Case vbNo
        End Select
    End Sub

    Private Sub GetListFilesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetListFilesToolStripMenuItem1.Click
        GLV.Show()
        Me.Hide()
    End Sub

    Private Sub RenameFilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RenameFilesToolStripMenuItem.Click
        RenameFiles.Show()
        Me.Hide()
    End Sub

    Private Sub FindReplace_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub
End Class