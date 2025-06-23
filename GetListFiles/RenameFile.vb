Imports System.IO

Public Class RenameFile
    Dim oldFile As String

    Private Sub RenameFile_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        TextBox2.Focus()
        TextBox1.Text = GLV.DataGridView1.CurrentCell.Value.ToString
        TextBox2.Text = GLV.DataGridView1.CurrentCell.Value.ToString
        Dim rowIndex As Integer = GLV.DataGridView1.CurrentCell.RowIndex
        Dim colIndex As Integer
        If GLV.cbFormSelValue = "WELL_FILE" Then
            colIndex = GLV.DataGridView1.ColumnCount - 3
        ElseIf GLV.cbFormSelValue = "WELL_LOG_DATA" Then
            colIndex = 18 - 1
        ElseIf GLV.cbFormSelValue = "WELL_LOG_IMAGE" Then
            colIndex = 23 - 1
        End If
        oldFile = GLV.DataGridView1.Rows(rowIndex).Cells(colIndex).Value.ToString & "\" & TextBox1.Text
    End Sub

    Private Sub TextBox2_Focus(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox2.Enter
        If TextBox2.Text.Length <> 0 Then
            btnSave.Enabled = True
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim respon As String = MsgBox("Are you sure you want to save?", MsgBoxStyle.YesNo, "CONFIRMATION")
        Dim fullPath As String
        Dim strPath As String
        Dim wellName As String = String.Empty

        Select Case respon
            Case vbYes
                Try
                    'Dim files() As String = Directory.GetFiles(GLV.txBrowse.Text, GLV.txFind.Text, SearchOption.AllDirectories)
                    'For Each oldFile As String In files
                    Dim oldFilePath As String = Path.GetFileName(oldFile).ToString
                    If oldFilePath.ToUpper = TextBox1.Text Then
                        fullPath = Path.GetFullPath(oldFile).ToString
                        strPath = Path.GetDirectoryName(oldFile.ToString)

                        If File.Exists(strPath + "\" + TextBox2.Text) Then
                            MessageBox.Show("File with the same filename has already existed", "FAILED", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Else
                            'If (TextBox2.Text.IndexOf("_") <> -1) Then
                            '    wellName = TextBox2.Text.Substring(0, TextBox2.Text.IndexOf("_"))
                            'End If

                            GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(1 - 1).Value = TextBox2.Text
                            If GLV.cbFormSelValue = "WELL_FILE" Then
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(2 - 1).Value = GLV.getWell(TextBox2.Text)

                                If GLV.fString = 2 Then
                                    GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(2 - 1).Style.ForeColor = Color.Red
                                ElseIf GLV.fString = 0 Then
                                    GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(2 - 1).Style.ForeColor = Color.Black
                                End If
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(4 - 1).Value = GLV.getBarcode(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(5 - 1).Value = GLV.getTitle(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(7 - 1).Value = GLV.getDate(TextBox2.Text)

                            ElseIf GLV.cbFormSelValue = "WELL_LOG_DATA" Then
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(22 - 1).Value = GLV.getWell(TextBox2.Text)

                                If GLV.fString = 2 Then
                                    GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(22 - 1).Style.ForeColor = Color.Red
                                ElseIf GLV.fString = 0 Then
                                    GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(22 - 1).Style.ForeColor = Color.Black
                                End If
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(13 - 1).Value = GLV.getRunNo(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(14 - 1).Value = GLV.getDate(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(15 - 1).Value = GLV.getCtLas(TextBox2.Text)

                            ElseIf GLV.cbFormSelValue = "WELL_LOG_IMAGE" Then
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(27 - 1).Value = GLV.getWell(TextBox2.Text)

                                If GLV.fString = 2 Then
                                    GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(27 - 1).Style.ForeColor = Color.Red
                                ElseIf GLV.fString = 0 Then
                                    GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(27 - 1).Style.ForeColor = Color.Black
                                End If
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(9 - 1).Value = GLV.getScale(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(14 - 1).Value = GLV.getBarcode(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(30 - 1).Value = GLV.getRunNo(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(31 - 1).Value = GLV.getDate(TextBox2.Text)
                                GLV.DataGridView1.Rows(GLV.DataGridView1.CurrentCell.RowIndex).Cells(36 - 1).Value = GLV.getCtTif(TextBox2.Text)
                            End If
                            File.Move(strPath + "\" + TextBox1.Text, strPath + "\" + TextBox2.Text)
                            MessageBox.Show("Filename has been successfuly changed", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Me.Close()
                        End If
                    End If
                    'Next
                Catch ex As Exception
                    MessageBox.Show(ex.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub
End Class