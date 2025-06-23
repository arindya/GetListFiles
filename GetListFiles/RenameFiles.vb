'+-----------------------------+
'|  Created By   : R.Armayndo  |
'|  Created Date : 20121106    |
'+-----------------------------+
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class RenameFiles
    Dim filelogErr As String

    Dim MyConnection As System.Data.OleDb.OleDbConnection
    Dim ExcelDataSet As System.Data.DataSet
    Dim ExcelAdapter As System.Data.OleDb.OleDbDataAdapter
    Dim stFilePathAndName As String
    Private Sub RenameFiles_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Dim stFileName As String


        OpenFileDialog1.InitialDirectory = System.Environment.CurrentDirectory
        OpenFileDialog1.Title = "Open xls File"
        OpenFileDialog1.Filter = "Text files (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = True

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            stFilePathAndName = OpenFileDialog1.FileName

            Dim MyFile As FileInfo = New FileInfo(stFilePathAndName)
            stFileName = MyFile.Name

            txBrowse.Text = stFileName
        End If


    End Sub

    Private Sub btnProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProses.Click
        ''''modul tulis file > START
        Dim bln As String = Date.Now.Month.ToString
        Dim tgll As String = Date.Now.Day.ToString

        If Len(bln) = 1 Then
            bln = "0" + bln
        End If
        If Len(tgll) = 1 Then
            tgll = "0" + tgll
        End If
        filelogErr = AppDomain.CurrentDomain.BaseDirectory + "\" + "RenameFilesLog" + Date.Now.Year.ToString + bln + tgll + "_error.txt"
        
        If File.Exists(filelogErr) Then
        Else
            Using sw As StreamWriter = File.CreateText(filelogErr)
                sw.WriteLine("::::Patra Nusa Data::::")
                sw.WriteLine(".......................")
                sw.WriteLine("--------------------------------------------------------------------------------------------------------------------------------")
                sw.WriteLine("-->>{0}", DateTime.Now.ToLongDateString())
                sw.WriteLine("--------------------------------------------------------------------------------------------------------------------------------")
                sw.WriteLine("")
                sw.WriteLine("")
                sw.Close()
            End Using
        End If
        ''''modul tulis file > END
     


        Dim strPath, namabaru, namalama, StrMsg As String
        Dim CountErr As Int16 = 0
        Dim CountSama As Int16 = 0
        Dim CountOK As Int16 = 0

        Dim jumbar As Integer = DataGridView1.RowCount
        Dim jumKol As Integer = DataGridView1.ColumnCount
        log(0, "total file", jumbar)

        Dim x As Integer = 0
        If jumKol > 3 Then
            MsgBox("Kolom > 3, hanya kolom Path, Nama_awal dan Nama_akhir yang di perbolehkan", MsgBoxStyle.Exclamation, "Perhatian")
        ElseIf jumKol = 3 And cbData.Checked = True Then
            ' Set up the progress bar's properties
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = jumbar
            ProgressBar1.Value = 0
            For x = 0 To jumbar - 1
                ProgressBar1.Value = ProgressBar1.Value + 1
                lbl_record.Text = x + 1 & " of " & jumbar.ToString & " Files"
                lbl_record.Refresh()
                'MsgBox(DataGridView1.Item(0, x).Value)
                System.Threading.Thread.Sleep(100)
                strPath = DataGridView1.Item(0, x).Value
                namalama = DataGridView1.Item(1, x).Value
                namabaru = DataGridView1.Item(2, x).Value

                If File.Exists(strPath + "\" + namalama) Then
                    ''File ada
                    If File.Exists(strPath + "\" + namabaru) Then
                        'do nothing
                        CountSama += 1
                        log(CountSama, "[FileExist]", namabaru)
                    Else
                        Try
                            If cbDelOri.Checked = True Then
                                File.Move(strPath + "\" + namalama, strPath + "\" + namabaru)
                            Else
                                File.Copy(strPath + "\" + namalama, strPath + "\" + namabaru)
                            End If

                            CountOK += 1
                        Catch ex As Exception
                            CountErr += 1
                            'MsgBox(ex.ToString, MsgBoxStyle.Critical, "ERROR")
                            log(CountErr, "[  ERROR  ]", ex.Message)
                        End Try
                    End If
                Else
                    ''File tidak ada
                    CountErr += 1
                    log(CountErr, "[Not Found]", namalama)
                End If

            Next
            StrMsg = "Finished, " + CountOK.ToString + " Files has been renamed."
            If CountSama <> 0 Then
                StrMsg = StrMsg + vbCrLf + CountSama.ToString + " Files cannot be renamed because destination name been exist"
            End If
            If CountErr <> 0 Then
                StrMsg = StrMsg + vbCrLf + CountErr.ToString + " Files cannot be renamed because unexpected ERROR"
            End If
            StrMsg = StrMsg + vbCrLf + " Log can be found in " + filelogErr

            MsgBox(StrMsg, MsgBoxStyle.Information, "FINISH")
        ElseIf jumKol > 2 Then
            MsgBox("Kolom > 2, hanya kolom Nama_awal dan Nama_akhir yang di perbolehkan", MsgBoxStyle.Exclamation, "Perhatian")
        ElseIf jumKol = 2 And txFolder.Text.Trim <> "" Then
            ' Set up the progress bar's properties
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = jumbar
            ProgressBar1.Value = 0
            For x = 0 To jumbar - 1
                ProgressBar1.Value = ProgressBar1.Value + 1
                lbl_record.Text = x + 1 & " of " & jumbar.ToString & " Files"
                lbl_record.Refresh()
                'MsgBox(DataGridView1.Item(0, x).Value)
                System.Threading.Thread.Sleep(100)
                namalama = DataGridView1.Item(0, x).Value
                namabaru = DataGridView1.Item(1, x).Value

                If File.Exists(txFolder.Text + "\" + namalama) Then
                    ''File ada
                    If File.Exists(txFolder.Text + "\" + namabaru) Then
                        'do nothing
                        CountSama += 1
                        log(CountSama, "[FileExist]", namabaru)
                    Else
                        Try
                            If cbDelOri.Checked = True Then
                                File.Move(txFolder.Text + "\" + namalama, txFolder.Text + "\" + namabaru)
                            Else
                                File.Copy(txFolder.Text + "\" + namalama, txFolder.Text + "\" + namabaru)
                            End If

                            CountOK += 1
                        Catch ex As Exception
                            CountErr += 1
                            'MsgBox(ex.ToString, MsgBoxStyle.Critical, "ERROR")
                            log(CountErr, "[  ERROR  ]", ex.Message)
                        End Try
                    End If
                Else
                    ''File tidak ada
                    CountErr += 1
                    log(CountErr, "[Not Found]", namalama)
                End If

            Next
            StrMsg = "Finished, " + CountOK.ToString + " Files has been renamed."
            If CountSama <> 0 Then
                StrMsg = StrMsg + vbCrLf + CountSama.ToString + " Files cannot be renamed because destination name been exist"
            End If
            If CountErr <> 0 Then
                StrMsg = StrMsg + vbCrLf + CountErr.ToString + " Files cannot be renamed because unexpected ERROR"
            End If
            StrMsg = StrMsg + vbCrLf + " Log can be found in " + filelogErr

            MsgBox(StrMsg, MsgBoxStyle.Information, "FINISH")
        Else
            MsgBox("Cek Input", MsgBoxStyle.Exclamation, "Perhatian")
        End If

        'MsgBox(DataGridView1.Item(0, 0).Value)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; " + _
                                "Data Source=" + stFilePathAndName + ";Extended Properties=Excel 12.0;")
        Try
            ExcelAdapter = New System.Data.OleDb.OleDbDataAdapter("select * from [" + txSheet.Text + "$" + txRange.Text + "]", MyConnection)
            ExcelAdapter.TableMappings.Add("Table", "Excel Data")
            ExcelDataSet = New System.Data.DataSet
            ExcelAdapter.Fill(ExcelDataSet)
            DataGridView1.DataSource = ExcelDataSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.ToString, "Importing Excel", _
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        MsgBox("load successfully with " + DataGridView1.RowCount.ToString + " rows", MsgBoxStyle.Information, "Success")
    End Sub

    Private Sub btnBrowse2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse2.Click
        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        ' Description that displays above the dialog box control.

        MyFolderBrowser.Description = "Select the Folder"
        MyFolderBrowser.ShowNewFolderButton = False
        ' Sets the root folder where the browsing starts from
        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txFolder.Text = MyFolderBrowser.SelectedPath
        End If
    End Sub
    Public Sub log(ByVal err_s As Integer, ByVal err_num As String, ByVal logMessage As String)
        Using sw As StreamWriter = File.AppendText(filelogErr)
            sw.Write("{0}; {1}", err_s, DateTime.Now.ToLongTimeString())
            sw.WriteLine(";{0};{1}", err_num, logMessage)
            sw.Flush()
            sw.Close()
        End Using
    End Sub

    Private Sub cbData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbData.CheckedChanged
        If cbData.Checked = True Then
            txFolder.Text = ""
            txFolder.Enabled = False
        Else
            txFolder.Enabled = True
        End If
    End Sub

    Private Sub GetListFilesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetListFilesToolStripMenuItem1.Click
        GLV.Show()
        Me.Hide()
    End Sub

    Private Sub FindReplaceFillesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindReplaceFillesToolStripMenuItem.Click
        FindReplace.Show()
        Me.Hide()
    End Sub

    Private Sub RenameFiles_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub
End Class
