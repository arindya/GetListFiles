'+-----------------------------+
'|  Created By   : R.Armayndo  |
'|  Created Date : 20120816    |
'+-----------------------------+
Imports System.IO
Imports System.Data.OleDb
Public Class Form1
    Dim files() As String

    Dim odbcn As OleDbConnection
    Dim odbcmd As OleDbCommand
    Dim odbda As OleDbDataAdapter
    Dim dt As Data.DataTable
    Dim objDataSet As New DataSet
    Dim sql_cek, pathdir As String
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

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txBrowse.Text = Application.StartupPath.ToString
    End Sub
    Public Function CheckFile(ByVal NamaFile As String) As Boolean
        Dim hasil As Boolean = True
        Dim jum_rec As Integer
        Dim koneksi As String = "Provider=OraOLEDB.Oracle;Data Source=" + txDB.Text.Trim + ";User ID=" + txID.Text.Trim + ";Password=" + txPass.Text.Trim
        Try
            objDataSet.Reset()
            odbcmd = New OleDbCommand
            odbcn = New OleDbConnection(koneksi)
            odbcn.Open()
            sql_cek = "select * from mcv_list_files where pathfile like '%" + NamaFile.ToUpper + "'"
            odbcmd = New OleDbCommand(sql_cek, odbcn)
            odbda = New OleDbDataAdapter(odbcmd)
            odbda.Fill(objDataSet, "dataFile")
            jum_rec = objDataSet.Tables("dataFile").Rows.Count
            If jum_rec > 0 Then
                hasil = False
            End If
            odbcn.Close()
            Return hasil
        Catch ex As Exception
            odbcn.Close()
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "ERROR")
            Me.Close()
        End Try
    End Function
    Private Sub btnProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess.Click
        Try
            If Directory.Exists(txBrowse.Text) And Directory.Exists(Directory.GetDirectoryRoot(txMove.Text)) Then
                TextBox1.Text = ""
                Dim tempString As String
                Dim jum, xcount As Integer
                Dim types() As String = Split(txType.Text, ";")
                'Dim cari As String = Trim(txFind.Text)
                files = Directory.GetFiles(txBrowse.Text, txFind.Text, SearchOption.AllDirectories)
                'files = Directory.GetFileSystemEntries(txBrowse.Text, "*")
                For Each filee As String In files
                    For Each type As String In types
                        'If filee.ToUpper.Contains(Trim(type)) And Trim(type) <> "" Then
                        If filee.ToUpper.EndsWith(Trim(type)) And Trim(type) <> "" Then
                            jum += 1
                            xcount += 1
                            System.Threading.Thread.Sleep(100)
                            'TextBox1.Text = Path.GetFileName(files(x).ToString) + vbCrLf + TextBox1.Text

                            TextBox1.Text = filee.ToString + vbCrLf + TextBox1.Text
                            TextBox1.Refresh()
                            Me.Refresh()

                            If xcount > 30 Then
                                'TextBox2.Text = TextBox2.Text.Substring(0, TextBox2.Text.Length / 3)
                                If (TextBox1.Text.Length) < 2501 Then
                                    TextBox1.Text = TextBox1.Text.Substring(0, TextBox1.Text.Length / 1.5)
                                Else
                                    TextBox1.Text = TextBox1.Text.Substring(0, 2500)
                                End If
                                xcount = 0
                            End If

                            tempString = filee.Replace(txDigDat.Text, "DigDat")
                            If CheckFile(tempString.ToString) Then
                                'MsgBox("TIDAK ADA" + tempString.ToString, MsgBoxStyle.Critical, "ERROR")
                                pathdir = Path.GetDirectoryName(txMove.Text + "\" + tempString)
                                If Not Directory.Exists(pathdir) Then
                                    Directory.CreateDirectory(pathdir)
                                End If
                                If Not File.Exists(txMove.Text + "\" + tempString) Then
                                    File.Move(filee.ToString, txMove.Text + "\" + tempString)
                                End If
                                Exit For
                            Else
                                'MsgBox("file ada", MsgBoxStyle.Critical, "ERROR")
                                Exit For
                            End If
                        End If
                    Next
                    'If file.ToUpper.Contains(".JPG") Or file.ToUpper.Contains(".JPEG") Or + _
                    'file.ToUpper.Contains(".TIF") Or file.ToUpper.Contains(".TIFF") Or file.ToUpper.Contains(".GIF") Or + _
                    'file.ToUpper.Contains(".PDF") Or file.ToUpper.Contains(".PDS") Or file.ToUpper.Contains(".LAS") Then
                    '    jum += 1
                    '    System.Threading.Thread.Sleep(100)
                    '    'TextBox1.Text = Path.GetFileName(files(x).ToString) + vbCrLf + TextBox1.Text
                    '    TextBox1.Text = file.ToString + vbCrLf + TextBox1.Text
                    '    TextBox1.Refresh()
                    'End If
                Next
                'Dim x As Integer
                'For x = 0 To files.Length - 1
                '    If files(x).Contains(".JPG") Then
                '        System.Threading.Thread.Sleep(100)
                '        TextBox1.Text = Path.GetFileName(files(x).ToString) + vbCrLf + TextBox1.Text
                '        'TextBox1.Text = files(x).ToString + vbCrLf + TextBox1.Text
                '        TextBox1.Refresh()
                '    End If
                'Next
                MsgBox("Searching complete, " + jum.ToString + " Files found", MsgBoxStyle.Information, "FINISH")
            Else
                MsgBox("One of directory does not exist", MsgBoxStyle.Exclamation, "WARNING")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class
