'+----------------------------------+
'|  Created By   : R.Armayndo       |
'|  Created Date : 20140822         |
'|  Modified By  : Pradhipta RH     |
'|  Modified Date: 2016/07/11       |
'|  Modified By  : M.Sultan AlFarid |
'|  Modified Date: 2025/06/12       |
'+----------------------------------+
Imports System.Threading
Imports System.Globalization
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports Oracle.ManagedDataAccess.Client
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Linq

Public Class GLV
    'Delegate Control
    Private Delegate Sub ToolStripStatusLabelDelegate1(ByVal value As String)
    Private Delegate Sub ToolStripStatusLabelDelegate2(ByVal value As String)
    Private Delegate Sub ToolStripProgressBarDelegate(ByVal value As Integer)
    Private Delegate Sub Del_BtnBrowse(ByVal value As Boolean)
    Private Delegate Sub Del_BtnBrowseFile(ByVal value As Boolean)
    Private Delegate Sub Del_BtnExecute(ByVal value As Boolean)
    Private Delegate Sub Del_BtnRename(ByVal value As Boolean)
    Private Delegate Sub Del_BtnExport(ByVal value As Boolean)
    Private Delegate Sub Del_BtnUpload(ByVal value As Boolean)
    Private Delegate Sub Del_BtnCancel(ByVal value As Boolean)
    Dim replaceAll As Boolean = False
    Dim skipAll As Boolean = False

    Private Sub ToolStripStatusLabelTxt1(ByVal value As String)
        If StatusStrip1.InvokeRequired Then
            Dim del As New ToolStripStatusLabelDelegate1(AddressOf ToolStripStatusLabelTxt1)
            StatusStrip1.Invoke(del, value)
        Else
            ToolStripStatusLabel1.Text = value
        End If
    End Sub

    Private Sub ToolStripStatusLabelTxt2(ByVal value As String)
        If StatusStrip1.InvokeRequired Then
            Dim del As New ToolStripStatusLabelDelegate2(AddressOf ToolStripStatusLabelTxt2)
            StatusStrip1.Invoke(del, value)
        Else
            ToolStripStatusLabel2.Text = value
        End If
    End Sub

    Private Sub ToolStripProgressBar(ByVal value As Integer)
        If StatusStrip1.InvokeRequired Then
            Dim del As New ToolStripProgressBarDelegate(AddressOf ToolStripProgressBar)
            StatusStrip1.Invoke(del, value)
        Else
            ToolStripProgressBar1.Value = value
        End If
    End Sub

    Private Sub _Del_BtnBrowse(ByVal value As Boolean)
        If btnExecute.InvokeRequired Then
            Dim del As New Del_BtnBrowse(AddressOf _Del_BtnBrowse)
            btnExecute.Invoke(del, value)
        Else
            btnBrowse.Enabled = value
        End If
    End Sub

    Private Sub _Del_BtnBrowseFile(ByVal value As Boolean)
        If btnExecute.InvokeRequired Then
            Dim del As New Del_BtnBrowse(AddressOf _Del_BtnBrowseFile)
            btnExecute.Invoke(del, value)
        Else
            btnBrowseFile.Enabled = value
        End If
    End Sub

    Private Sub _Del_BtnExecute(ByVal value As Boolean)
        If btnExecute.InvokeRequired Then
            Dim del As New Del_BtnExecute(AddressOf _Del_BtnExecute)
            btnExecute.Invoke(del, value)
        Else
            btnExecute.Enabled = value
        End If
    End Sub

    Private Sub _Del_BtnRename(ByVal value As Boolean)
        If btnExecute.InvokeRequired Then
            Dim del As New Del_BtnRename(AddressOf _Del_BtnRename)
            btnExecute.Invoke(del, value)
        Else
            btnRename.Enabled = value
        End If
    End Sub

    Private Sub _Del_BtnExport(ByVal value As Boolean)
        If btnExecute.InvokeRequired Then
            Dim del As New Del_BtnExport(AddressOf _Del_BtnExport)
            btnExport.Invoke(del, value)
        Else
            btnExport.Enabled = value
        End If
    End Sub

    Private Sub _Del_BtnUpload(ByVal value As Boolean)
        If btnExecute.InvokeRequired Then
            Dim del As New Del_BtnUpload(AddressOf _Del_BtnUpload)
            btnUpload.Invoke(del, value)
        Else
            btnUpload.Enabled = value
        End If
    End Sub

    Private Sub _Del_BtnCancel(ByVal value As Boolean)
        If btnCancel.InvokeRequired Then
            Dim del As New Del_BtnCancel(AddressOf _Del_BtnCancel)
            btnCancel.Invoke(del, value)
        Else
            btnCancel.Enabled = value
        End If
    End Sub

    Dim WithEvents dgv As System.Windows.Forms.DataGridView
    Dim files() As String
    Dim respon As String
    Dim bln As String = Date.Now.Month.ToString
    Dim tgll As String = Date.Now.Day.ToString
    Dim filelogErr As String
    Dim xcount As Int16 = 0
    Dim CountErr As Integer
    Public fString As Int16 'fString : 0 = OK | 1 = Warning | 2 = Error
    Dim dataGroup As String
    Public formType As String
    Dim browseStatus As Integer
    Dim activeCell As DataGridViewCell
    Dim newCulture As CultureInfo
    Dim OldCulture As CultureInfo

    Dim uploadErr As String = String.Empty
    Dim copyCntErr As Integer
    Dim copyFlag As Integer
    Dim insertCntErr As Integer
    Dim insertFlag As Integer
    Public cbFormSelValue As String
    Dim successFlg As Integer

    Dim odbcn As OracleConnection
    Dim odbcmd As OracleCommand
    Dim odbda As OracleDataAdapter
    Dim objDataSet As New DataSet
    Dim sql_cek As String

    Dim conString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"
    Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"
    Dim query As String
    Dim query1 As String
    Dim dataAdp As OracleDataAdapter
    Dim dataSet As DataSet = New DataSet()


    Private Sub btnBrowse_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnBrowse.Click
        txBrowse.Text = String.Empty

        Dim MyFolderBrowser As New FolderBrowserDialog
        ' Description that displays above the dialog box control.

        MyFolderBrowser.Description = "Select the Folder"
        MyFolderBrowser.ShowNewFolderButton = False

        ' Sets the root folder where the browsing starts from
        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = DialogResult.OK Then
            txBrowse.Text = MyFolderBrowser.SelectedPath
            txBrowseFile.Text = String.Empty
            cbForm.Enabled = True
            browseStatus = 1
        End If
    End Sub

    Private Sub btnBrowseFile_Click(sender As Object, e As EventArgs) Handles btnBrowseFile.Click
        txBrowseFile.Text = String.Empty

        Dim FileBrowse As New OpenFileDialog

        FileBrowse.Title = "Open File Dialog"
        FileBrowse.Filter = "Excel 2003 (*.xls)|*.xls|Excel 2007 (*.xlsx)|*.xlsx"
        FileBrowse.RestoreDirectory = True
        FileBrowse.Multiselect = False

        If FileBrowse.ShowDialog() = DialogResult.OK Then
            txBrowseFile.Text = FileBrowse.FileName
            txBrowse.Text = String.Empty
            cbForm.Enabled = True
            browseStatus = 2
        End If

    End Sub

    Private Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExecute.Click
        _Del_BtnBrowse(False)
        _Del_BtnBrowseFile(False)
        _Del_BtnExecute(False)
        _Del_BtnExport(False)
        _Del_BtnUpload(False)
        'MessageBox.Show("FTP User: " & My.Settings.FTPUser & vbCrLf &
        '        "Digdat Host: " & My.Settings.DigdatHost, "Konfigurasi FTP")


        CountErr = 0
        cbFormSelValue = cbForm.SelectedValue.ToString
        Dim strCompare As String = String.Empty

        If browseStatus = 1 Then
            If Trim(txBrowse.Text) <> "" Then
                If Directory.Exists(txBrowse.Text) Then

                    Dim types() As String = Split(txType.Text, ";")

                    If cbFindMod.Checked = True Then
                        files = Directory.GetFiles(txBrowse.Text, txFind.Text, SearchOption.AllDirectories)
                    Else
                        files = Directory.GetFiles(txBrowse.Text, txFind.Text, SearchOption.TopDirectoryOnly)
                    End If

                    Array.Sort(files)

                    Dim x As Integer

                    respon = MessageBox.Show(files.Length.ToString + " file(s) found, Do you want to continue ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    Select Case respon
                        Case vbYes

                            Dim xlApp As Excel.Application = New Excel.Application()
                            Dim xlWorkBook As Excel.Workbook
                            Dim xlWorkSheet As Excel.Worksheet
                            Dim misValue As Object = Reflection.Missing.Value
                            Dim dateStart As Date = Date.Now

                            Dim oldci As CultureInfo = Thread.CurrentThread.CurrentCulture
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-us")

                            If xlApp Is Nothing Then
                                MessageBox.Show("Excel is not properly installed!!")
                                Return
                            End If

                            If cbForm.SelectedIndex = 0 Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "Path"
                                xlWorkSheet.Cells(1, 2) = "File Name"
                            ElseIf cbForm.SelectedValue.ToString = "WELL_FILE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WF_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 3) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 4) = "WF_BARCODE"
                                xlWorkSheet.Cells(1, 5) = "WF_TITLE"
                                xlWorkSheet.Cells(1, 6) = "WF_AUTHORS"
                                xlWorkSheet.Cells(1, 7) = "WF_DATE"
                                xlWorkSheet.Cells(1, 8) = "WF_TYPE"
                                xlWorkSheet.Cells(1, 9) = "WF_SUBJECT"
                                xlWorkSheet.Cells(1, 10) = "WF_GROUP"
                                xlWorkSheet.Cells(1, 11) = "WF_NUM_OF_PAGE"
                                xlWorkSheet.Cells(1, 12) = "WF_NOTE"
                                xlWorkSheet.Cells(1, 13) = "WF_DOC_VER"
                                xlWorkSheet.Cells(1, 14) = "WF_FILE_PATH"
                                xlWorkSheet.Cells(1, 15) = "WF_FILE_SIZE"
                                xlWorkSheet.Cells(1, 16) = "WF_LOAD_BY"
                                xlWorkSheet.Cells(1, 17) = "WF_LOADED_DATE"
                                xlWorkSheet.Cells(1, 18) = "WF_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 19) = "WF_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 24) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 25) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 26) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 27) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 28) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 29) = "FLAG"
                                xlWorkSheet.Cells(1, 30) = "FLAG_UPLOAD"
                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "S1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_DATA" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WLD_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WLD_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WLD_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WLD_NOTE"
                                xlWorkSheet.Cells(1, 5) = "WLD_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 6) = "WLD_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 7) = "WLD_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "WLD_LOAD_DATE"
                                xlWorkSheet.Cells(1, 9) = "WELL_LOG_S"
                                xlWorkSheet.Cells(1, 10) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 11) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 12) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 18) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 19) = "FLAG"
                                xlWorkSheet.Cells(1, 20) = "FLAG_UPLOAD"
                                xlWorkSheet.Cells(1, 21) = "WL_PRODUCERS"
                                xlWorkSheet.Cells(1, 22) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 23) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 24) = "WL_LOG_TYPE"
                                xlWorkSheet.Cells(1, 25) = "WL_RUN_NO"
                                xlWorkSheet.Cells(1, 26) = "WL_RUN_DATE"
                                xlWorkSheet.Cells(1, 27) = "WL_TOP_DEPTH"
                                xlWorkSheet.Cells(1, 28) = "WL_BOTTOM_DEPTH"
                                xlWorkSheet.Cells(1, 29) = "WL_DEPTH_U"
                                xlWorkSheet.Cells(1, 30) = "WL_REMARKS"
                                xlWorkSheet.Cells(1, 31) = "WL_CURVE_TYPE"
                                xlWorkSheet.Cells(1, 32) = "WL_NOTE"
                                xlWorkSheet.Cells(1, 33) = "WL_TITLE"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_IMAGE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WLI_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WLI_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WLI_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WLI_NOTE"
                                xlWorkSheet.Cells(1, 5) = "WLI_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 6) = "WLI_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 7) = "WLI_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "WLI_LOAD_DATE"
                                xlWorkSheet.Cells(1, 9) = "WLI_VERTICAL_SCALE"
                                xlWorkSheet.Cells(1, 10) = "WELL_LOG_S"
                                xlWorkSheet.Cells(1, 11) = "WLI_HDR_FILE_NAME"
                                xlWorkSheet.Cells(1, 12) = "WLI_HDR_FILE_PATH"
                                xlWorkSheet.Cells(1, 13) = "WLI_HDR_FILE_SIZE"
                                xlWorkSheet.Cells(1, 14) = "WLI_BARCODE"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 23) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 24) = "FLAG"
                                xlWorkSheet.Cells(1, 25) = "FLAG_UPLOAD"
                                xlWorkSheet.Cells(1, 26) = "WL_PRODUCERS"
                                xlWorkSheet.Cells(1, 27) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 28) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 29) = "WL_LOG_TYPE"
                                xlWorkSheet.Cells(1, 30) = "WL_RUN_NO"
                                xlWorkSheet.Cells(1, 31) = "WL_RUN_DATE"
                                xlWorkSheet.Cells(1, 32) = "WL_TOP_DEPTH"
                                xlWorkSheet.Cells(1, 33) = "WL_BOTTOM_DEPTH"
                                xlWorkSheet.Cells(1, 34) = "WL_DEPTH_U"
                                xlWorkSheet.Cells(1, 35) = "WL_REMARKS"
                                xlWorkSheet.Cells(1, 36) = "WL_CURVE_TYPE"
                                xlWorkSheet.Cells(1, 37) = "WL_NOTE"
                                xlWorkSheet.Cells(1, 38) = "WL_TITLE"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "T1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_MASTER_LOG" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WML_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WML_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WML_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WML_TITLE"
                                xlWorkSheet.Cells(1, 5) = "WML_NOTE"
                                xlWorkSheet.Cells(1, 6) = "WML_DATE"
                                xlWorkSheet.Cells(1, 7) = "WML_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 8) = "WML_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 9) = "WML_LOAD_BY"
                                xlWorkSheet.Cells(1, 10) = "WML_DOC_VER"
                                xlWorkSheet.Cells(1, 11) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 12) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 13) = "WML_LOAD_DATE"
                                xlWorkSheet.Cells(1, 14) = "WML_BARCODE"
                                xlWorkSheet.Cells(1, 15) = "WML_VERTICAL_SCALE"
                                xlWorkSheet.Cells(1, 16) = "WML_TOP_DEPTH"
                                xlWorkSheet.Cells(1, 17) = "WML_BOTTOM_DEPTH"
                                xlWorkSheet.Cells(1, 18) = "WML_DEPTH_U"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 24) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 25) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 26) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 27) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 28) = "FLAG"
                                xlWorkSheet.Cells(1, 29) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_CORRELATION" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WC_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WC_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WC_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WC_LOAD_BY"
                                xlWorkSheet.Cells(1, 5) = "WC_LOAD_DATE"
                                xlWorkSheet.Cells(1, 6) = "WC_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 7) = "WC_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 8) = "WC_NOTE"
                                xlWorkSheet.Cells(1, 9) = "WC_STATUS"
                                xlWorkSheet.Cells(1, 10) = "STRUCTURE_S"
                                xlWorkSheet.Cells(1, 11) = "WC_TITLE"
                                xlWorkSheet.Cells(1, 12) = "WC_DATE"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 21) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 22) = "FLAG"
                                xlWorkSheet.Cells(1, 23) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_HISTORY_PORFILE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 2) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 3) = "WHP_LAST_UPDATE"
                                xlWorkSheet.Cells(1, 4) = "WHP_TOTAL_DEPTH"
                                xlWorkSheet.Cells(1, 5) = "WHP_PLUG"
                                xlWorkSheet.Cells(1, 6) = "WHP_START_PERFO"
                                xlWorkSheet.Cells(1, 7) = "WHP_END_PERFO"
                                xlWorkSheet.Cells(1, 8) = "WHP_OPEN_HOLE"
                                xlWorkSheet.Cells(1, 9) = "WHP_GLV"
                                xlWorkSheet.Cells(1, 10) = "WHP_TOP_FISH"
                                xlWorkSheet.Cells(1, 11) = "WHP_UNIT"
                                xlWorkSheet.Cells(1, 12) = "WHP_PROD_LAYER"
                                xlWorkSheet.Cells(1, 13) = "WHP_PROD_STRING"
                                xlWorkSheet.Cells(1, 14) = "WHP_PUMP_STRING"
                                xlWorkSheet.Cells(1, 15) = "WHP_LIFTING"
                                xlWorkSheet.Cells(1, 16) = "WHP_IMG_FILE_NAME"
                                xlWorkSheet.Cells(1, 17) = "WHP_ACTIVITY"
                                xlWorkSheet.Cells(1, 18) = "WHP_SCANNED_FILE_NAME"
                                xlWorkSheet.Cells(1, 19) = "WHP_DOC_FILE_NAME"
                                xlWorkSheet.Cells(1, 20) = "WHP_WS_FILE_NAME"
                                xlWorkSheet.Cells(1, 21) = "WHP_CREATED_BY"
                                xlWorkSheet.Cells(1, 22) = "WHP_CREATED_DATE"
                                xlWorkSheet.Cells(1, 23) = "WHP_INTERVAL"
                                xlWorkSheet.Cells(1, 24) = "WHP_SETPACKER"
                                xlWorkSheet.Cells(1, 25) = "WHP_SROD_SIZE"
                                xlWorkSheet.Cells(1, 26) = "WHP_SROD_JTS"
                                xlWorkSheet.Cells(1, 27) = "WHP_TUBING_ID"
                                xlWorkSheet.Cells(1, 28) = "WHP_TUBING_OD"
                                xlWorkSheet.Cells(1, 29) = "WHP_TUBING_JTS"
                                xlWorkSheet.Cells(1, 30) = "WHP_PSN"
                                xlWorkSheet.Cells(1, 31) = "WHP_EOS"
                                xlWorkSheet.Cells(1, 32) = "VAL_WELL"
                                xlWorkSheet.Cells(1, 33) = "DESCRIPTION"
                                xlWorkSheet.Cells(1, 34) = "FILE_NAME"
                                xlWorkSheet.Cells(1, 35) = "FILE_SIZE"
                                xlWorkSheet.Cells(1, 35) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 36) = "FLAG"
                                xlWorkSheet.Cells(1, 37) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "GNG_CONTOUR_IMAGE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "GCI_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "GCI_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "GCI_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "GCI_NOTE"
                                xlWorkSheet.Cells(1, 5) = "GCI_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 6) = "GCI_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 7) = "GCI_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "GCI_LOAD_DATE"
                                xlWorkSheet.Cells(1, 9) = "GCI_STATUS"
                                xlWorkSheet.Cells(1, 10) = "STRUCTURE_S"
                                xlWorkSheet.Cells(1, 11) = "GCI_CONTOUR_TYPE"
                                xlWorkSheet.Cells(1, 12) = "GCI_TITLE"
                                xlWorkSheet.Cells(1, 13) = "GCI_DATE"
                                xlWorkSheet.Cells(1, 14) = "GCI_AUTHOR"
                                xlWorkSheet.Cells(1, 15) = "GCI_BARCODE"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 24) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 25) = "FLAG"
                                xlWorkSheet.Cells(1, 26) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "GNG_REPORT" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "GRI_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "GRI_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "GRI_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "GRI_NOTE"
                                xlWorkSheet.Cells(1, 5) = "GRI_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 6) = "GRI_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 7) = "GRI_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "GRI_LOAD_DATE"
                                xlWorkSheet.Cells(1, 9) = "GNG_REPORT_S"
                                xlWorkSheet.Cells(1, 10) = "GRI_DOC_VER"
                                xlWorkSheet.Cells(1, 11) = "GRI_SUBJECT"
                                xlWorkSheet.Cells(1, 12) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 20) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 21) = "FLAG"
                                xlWorkSheet.Cells(1, 22) = "FLAG_UPLOAD"
                                xlWorkSheet.Cells(1, 23) = "GR_TITLE"
                                xlWorkSheet.Cells(1, 24) = "GR_AUTHORS"
                                xlWorkSheet.Cells(1, 25) = "GR_DATE"
                                xlWorkSheet.Cells(1, 26) = "GR_TYPE"
                                xlWorkSheet.Cells(1, 27) = "GR_SUBJECT"
                                xlWorkSheet.Cells(1, 28) = "GR_NUM_OF_PAGE"
                                xlWorkSheet.Cells(1, 29) = "GR_NUM_ENCLOSURE"
                                xlWorkSheet.Cells(1, 30) = "GR_UPDATED_BY"
                                xlWorkSheet.Cells(1, 31) = "GR_UPDATED_DATE"
                                xlWorkSheet.Cells(1, 32) = "GR_NOTE"
                                xlWorkSheet.Cells(1, 33) = "STRUCTURE_S"
                                xlWorkSheet.Cells(1, 34) = "GR_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 35) = "GR_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 36) = "GR_BARCODE"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "REPO_WELL" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 2) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 3) = "RW_FTYPE"
                                xlWorkSheet.Cells(1, 4) = "RW_FCAT"
                                xlWorkSheet.Cells(1, 5) = "RW_FPATH"
                                xlWorkSheet.Cells(1, 6) = "RW_FNAME"
                                xlWorkSheet.Cells(1, 7) = "RW_FTITLE"
                                xlWorkSheet.Cells(1, 8) = "RW_FDATE"
                                xlWorkSheet.Cells(1, 9) = "RW_FSIZE"
                                xlWorkSheet.Cells(1, 10) = "RW_FSOURCE"
                                xlWorkSheet.Cells(1, 11) = "RW_LOAD_BY"
                                xlWorkSheet.Cells(1, 12) = "RW_LOADED_DATE"
                                xlWorkSheet.Cells(1, 13) = "RW_BARCODE"
                                xlWorkSheet.Cells(1, 14) = "RW_DESC"
                                xlWorkSheet.Cells(1, 15) = "RW_TAG"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 24) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 25) = "FLAG"
                                xlWorkSheet.Cells(1, 26) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "REPO_ASET" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "ASET_ID"
                                xlWorkSheet.Cells(1, 2) = "RA_FTYPE"
                                xlWorkSheet.Cells(1, 3) = "RA_FPATH"
                                xlWorkSheet.Cells(1, 4) = "RA_FNAME"
                                xlWorkSheet.Cells(1, 5) = "RA_FTITLE"
                                xlWorkSheet.Cells(1, 6) = "RA_FDATE"
                                xlWorkSheet.Cells(1, 7) = "RA_FSIZE"
                                xlWorkSheet.Cells(1, 8) = "RA_FSOURCE"
                                xlWorkSheet.Cells(1, 9) = "RA_LOAD_BY"
                                xlWorkSheet.Cells(1, 10) = "RA_LOADED_DATE"
                                xlWorkSheet.Cells(1, 11) = "RA_BARCODE"
                                xlWorkSheet.Cells(1, 12) = "RA_DESC"
                                xlWorkSheet.Cells(1, 13) = "RA_TAG"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 22) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 23) = "FLAG"
                                xlWorkSheet.Cells(1, 24) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "SITUATION_MAP_IMAGE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "SMI_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "SMI_FILE_PATH"
                                xlWorkSheet.Cells(1, 3) = "SMI_TITLE"
                                xlWorkSheet.Cells(1, 4) = "SMI_FILE_SIZE"
                                xlWorkSheet.Cells(1, 5) = "SMI_DATE"
                                xlWorkSheet.Cells(1, 6) = "SMI_LOAD_BY"
                                xlWorkSheet.Cells(1, 7) = "SMI_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 8) = "SMI_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 9) = "SMI_NOTE"
                                xlWorkSheet.Cells(1, 10) = "SMI_STATUS"
                                xlWorkSheet.Cells(1, 11) = "STRUCTURE_S"
                                xlWorkSheet.Cells(1, 12) = "SMI_FILE_DWG_NAME"
                                xlWorkSheet.Cells(1, 13) = "SMI_FILE_DWG_PATH"
                                xlWorkSheet.Cells(1, 14) = "SMI_FILE_DWG_SIZE"
                                xlWorkSheet.Cells(1, 15) = "SMI_MAP_SUBJECT"
                                xlWorkSheet.Cells(1, 16) = "SITUATION_MAP_IMAGE_S"
                                xlWorkSheet.Cells(1, 17) = "SMI_FILE_NAMES"
                                xlWorkSheet.Cells(1, 18) = "SMI_BARCODE"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 24) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 25) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 26) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 27) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 28) = "FLAG"
                                xlWorkSheet.Cells(1, 29) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "SP_DIGITAL_DATA" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "SPDD_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "SPDD_FILE_PATH"
                                xlWorkSheet.Cells(1, 3) = "SPDD_FILE_SIZE"
                                xlWorkSheet.Cells(1, 4) = "SP_ID"
                                xlWorkSheet.Cells(1, 5) = "SPDD_LOAD_BY"
                                xlWorkSheet.Cells(1, 6) = "SPDD_LOAD_DATE"
                                xlWorkSheet.Cells(1, 7) = "SPDD_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 8) = "SPDD_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 9) = "SPDD_NOTE"
                                xlWorkSheet.Cells(1, 10) = "SPDD_STATUS"
                                xlWorkSheet.Cells(1, 11) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 12) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 19) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 20) = "FLAG"
                                xlWorkSheet.Cells(1, 21) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "PROCESS_FACILITY_MAP" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "PFM_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "PFM_FILE_PATH"
                                xlWorkSheet.Cells(1, 3) = "PFM_FILE_SIZE"
                                xlWorkSheet.Cells(1, 4) = "PFM_TITLE"
                                xlWorkSheet.Cells(1, 5) = "PFM_SUBJECT"
                                xlWorkSheet.Cells(1, 6) = "PFM_AREA"
                                xlWorkSheet.Cells(1, 7) = "PFM_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "PFM_LOADED_DATE"
                                xlWorkSheet.Cells(1, 9) = "PFM_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 10) = "PFM_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 11) = "PFM_SCALE"
                                xlWorkSheet.Cells(1, 12) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 20) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 21) = "FLAG"
                                xlWorkSheet.Cells(1, 22) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "REGIONAL_MAP_IMAGE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "RMI_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "RMI_FILE_PATH"
                                xlWorkSheet.Cells(1, 3) = "RMI_TITLE"
                                xlWorkSheet.Cells(1, 4) = "RMI_FILE_SIZE"
                                xlWorkSheet.Cells(1, 5) = "RMI_DATE"
                                xlWorkSheet.Cells(1, 6) = "RMI_LOAD_BY"
                                xlWorkSheet.Cells(1, 7) = "RMI_NOTE"
                                xlWorkSheet.Cells(1, 8) = "RMI_STATUS"
                                xlWorkSheet.Cells(1, 9) = "RMI_AREA"
                                xlWorkSheet.Cells(1, 10) = "RMI_BARCODE"
                                xlWorkSheet.Cells(1, 11) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 12) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 19) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 20) = "FLAG"
                                xlWorkSheet.Cells(1, 21) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            Else
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                            End If

                            Array.Reverse(files)
                            CreateFile()
                            log(0, "File processed", (files.Length).ToString)

                            Dim rowXls As Integer = 1
                            For x = 0 To files.Length - 1
                                For Each typee As String In types
                                    Try
                                        If Path.GetFileName(files(x)).ToUpper.EndsWith(Trim(typee)) And Trim(typee) <> "" Then
                                            If Not IsNothing(xlApp) Then
                                                Dim fileinfo = New FileInfo(Path.GetFullPath(files(x)))
                                                Dim fileSize As Double = Math.Round(fileinfo.Length / 1024)

                                                If cbForm.SelectedIndex = 0 Then
                                                    xlWorkSheet.Cells(rowXls + 1, 1) = Path.GetDirectoryName(files(x).ToString.ToUpper)
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_FILE" Then
                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = getWell(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "SBS"
                                                    xlWorkSheet.Cells(rowXls + 1, 4) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 5) = getTitle(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 6) = getAuthor(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 7).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 7).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 7).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 8) = getWfType(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 8).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 8).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 9) = getWfSbj(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 9).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    If (Path.GetFileName(files(x).ToString.ToUpper.EndsWith(".PDF"))) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 11) = getNumberOfPdfPages(Path.GetFullPath(files(x).ToString)).ToString
                                                        If fString <> 0 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 11).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 11).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            fString = 0
                                                        End If
                                                    End If

                                                    'Check Filename to be used for Destination Path
                                                    dataGroup = xlWorkSheet.Cells(rowXls + 1, 1).Value.IndexOf("PETRAN")
                                                    If (dataGroup > -1) Then
                                                        dataGroup = "GNG"
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLLOG\PETRAN"
                                                    ElseIf xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString.ToUpper.Contains("PENGUKURAN") And xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString.ToUpper.Contains("TEKANAN") And xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString.ToUpper.Contains("DASAR") Then
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLTEST"
                                                    Else
                                                        dataGroup = "PETRO"
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLREPORT\WELLREPORTIMAGE"
                                                    End If

                                                    If (xlWorkSheet.Cells(rowXls + 1, 8).value = "PRE" And xlWorkSheet.Cells(rowXls + 1, 9).value = "B3") And (Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).value)) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLTEST"
                                                    ElseIf (xlWorkSheet.Cells(rowXls + 1, 8).value = "PPA" And xlWorkSheet.Cells(rowXls + 1, 9).value = "D4") And (Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).value)) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLLOG\PETRAN"
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 15) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 16) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = getApprovalGroupName(xlWorkSheet.Cells(rowXls + 1, 2).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, dataGroup)
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 28) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = 0

                                                    If Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).Value) Then
                                                        'Check PK in database
                                                        CheckPKWellFile(xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 2).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 8).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 9).Value.ToString)
                                                        If fString = 1 Then
                                                            xlWorkSheet.Cells(rowXls + 1, 29) = 1
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            xlWorkSheet.Cells(rowXls + 1, 8).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                        Else
                                                            xlWorkSheet.Cells(rowXls + 1, 29) = 0
                                                        End If
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 30) = 0

                                                    If Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 4).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 5).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 6).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).Value) And xlWorkSheet.Cells(rowXls + 1, 29).value = 0 Then
                                                        xlWorkSheet.Cells(rowXls + 1, 30) = 1
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_DATA" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\WELLLOG\WELLLOGDATA"
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 10) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                    xlWorkSheet.Cells(rowXls + 1, 18) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 19) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = 1
                                                    xlWorkSheet.Cells(rowXls + 1, 22) = wellname
                                                    xlWorkSheet.Cells(rowXls + 1, 23) = "SBS"

                                                    xlWorkSheet.Cells(rowXls + 1, 25) = getRunNo(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 25).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 25).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 25).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 26) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 26).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 26).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 26).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = getElevetionUnit(wellname)
                                                    xlWorkSheet.Cells(rowXls + 1, 31) = getCtLas(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        ' Format
                                                        xlWorkSheet.Cells(rowXls + 1, 31).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        fString = 0
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_IMAGE" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    If Path.GetFileName(files(x).ToString.ToUpper).Contains("_HDR.") Then
                                                        'If rowXls > 1 Then
                                                        rowXls -= 1
                                                        'End If

                                                        xlWorkSheet.Cells(rowXls + 1, 11) = Path.GetFileName(files(x).ToString.ToUpper)
                                                        xlWorkSheet.Cells(rowXls + 1, 12) = "DIGDAT\WELLLOG\WELLLOGIMAGE_HDR"
                                                        xlWorkSheet.Cells(rowXls + 1, 13) = fileSize.ToString
                                                        rowXls += 1
                                                        Continue For
                                                    Else
                                                        xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                        xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\WELLLOG\WELLLOGIMAGE"
                                                        xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                        xlWorkSheet.Cells(rowXls + 1, 9) = getScale(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 9).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 14).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 15) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                        xlWorkSheet.Cells(rowXls + 1, 23) = Path.GetDirectoryName(files(x).ToString)
                                                        xlWorkSheet.Cells(rowXls + 1, 24) = 0
                                                        xlWorkSheet.Cells(rowXls + 1, 25) = 1
                                                        xlWorkSheet.Cells(rowXls + 1, 27) = wellname
                                                        xlWorkSheet.Cells(rowXls + 1, 28) = "SBS"
                                                        xlWorkSheet.Cells(rowXls + 1, 30) = getRunNo(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 30).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 30).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 30).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 31) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 31).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 31).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 31).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If

                                                        xlWorkSheet.Cells(rowXls + 1, 34) = getElevetionUnit(wellname)

                                                        xlWorkSheet.Cells(rowXls + 1, 36) = getCtTif(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 36).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            fString = 0
                                                        End If

                                                        If Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 1).Value) Then
                                                            'Check PK in database
                                                            CheckPKWellLogImage(xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString)
                                                            If fString = 1 Then
                                                                xlWorkSheet.Cells(rowXls + 1, 24) = 1
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            End If
                                                        End If
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_MASTER_LOG" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\WELLLOG\MASTERLOG"
                                                    xlWorkSheet.Cells(rowXls + 1, 9) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 11) = wellname
                                                    xlWorkSheet.Cells(rowXls + 1, 12) = "SBS"

                                                    xlWorkSheet.Cells(rowXls + 1, 14) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 14).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 15) = getScale(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 18) = getElevetionUnit(wellname)
                                                    xlWorkSheet.Cells(rowXls + 1, 19) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                    xlWorkSheet.Cells(rowXls + 1, 27) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 28) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = 1

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_CORRELATION" Then
                                                    Dim structureName As String = getStructure(Path.GetFileName(files(x).ToString.ToUpper), 3)

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\STRGNGMAP\STRWELLCORRELATION"
                                                    xlWorkSheet.Cells(rowXls + 1, 4) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 10) = getStructure(Path.GetFileName(files(x).ToString.ToUpper), 2)
                                                    xlWorkSheet.Cells(rowXls + 1, 16) = getApprovalGroupName(structureName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 22) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 23) = 1

                                                ElseIf cbForm.SelectedValue.ToString = "GNG_CONTOUR_IMAGE" Then
                                                    Dim structureName As String = getStructure(Path.GetFileName(files(x).ToString.ToUpper), 3)

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\STRCONTOUR\STRCONTOURIMAGE"
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 10) = getStructure(Path.GetFileName(files(x).ToString.ToUpper), 2)

                                                    xlWorkSheet.Cells(rowXls + 1, 13) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 13).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 13).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 13).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 15) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 16) = getApprovalGroupName(structureName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 24) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 25) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 26) = 1


                                                ElseIf cbForm.SelectedValue.ToString = "GNG_REPORT" Then
                                                    Dim structureName As String = getStructure(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\STRGNGREPORT\STRGNGREPORTIMAGE"
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 12) = getApprovalGroupName(structureName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 22) = 1
                                                    xlWorkSheet.Cells(rowXls + 1, 23) = getTitle(Path.GetFileName(files(x).ToString.ToUpper), 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 25) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 28) = getNumberOfPdfPages(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 33) = getStructure(Path.GetFileName(files(x).ToString.ToUpper), 1)

                                                    xlWorkSheet.Cells(rowXls + 1, 36) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 36).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 36).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 36).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "REPO_WELL" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = wellname
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = "SBS"
                                                    'xlWorkSheet.Cells(rowXls + 1, 5) = "DIGDAT\STRGNGMAP\STRWELLCORRELATION"

                                                    xlWorkSheet.Cells(rowXls + 1, 6) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getTitle(Path.GetFileName(files(x).ToString.ToUpper), 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 8) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 9) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 11) = getLoadBy(Login.TextBox1.Text)

                                                    xlWorkSheet.Cells(rowXls + 1, 13) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 13).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 13).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 13).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 16) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                    xlWorkSheet.Cells(rowXls + 1, 24) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 25) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 26) = 1

                                                ElseIf cbForm.SelectedValue.ToString = "REPO_ASET" Then
                                                    Dim assetName As String = getAsset(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = assetName

                                                    If String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 2).Value) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 2) = getWfType(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                        xlWorkSheet.Cells(rowXls + 1, 24) = 0
                                                    Else
                                                        xlWorkSheet.Cells(rowXls + 1, 24) = 1
                                                    End If


                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\REPO\FIELD"

                                                    xlWorkSheet.Cells(rowXls + 1, 4) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 5) = getRATitle(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 6) = getRA_Fdate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 9) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 14) = getApprovalGroupName(assetName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 22) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 23) = 0

                                                ElseIf cbForm.SelectedValue.ToString = "SITUATION_MAP_IMAGE" Then

                                                    If Path.GetExtension(files(x).ToString.ToUpper) <> "DWG" Then
                                                        xlWorkSheet.Cells(rowXls + 1, 1) = Path.GetFileName(files(x).ToString.ToUpper)
                                                        xlWorkSheet.Cells(rowXls + 1, 2) = "DIGDAT\SFACILITY\SITUATIONMAP"
                                                        xlWorkSheet.Cells(rowXls + 1, 4) = fileSize.ToString
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 3) = getTitle(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 6) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 11) = getStructure(Path.GetFileName(files(x).ToString.ToUpper), 2)

                                                    If Path.GetExtension(files(x).ToString.ToUpper) = "DWG" Then
                                                        xlWorkSheet.Cells(rowXls + 1, 12) = Path.GetFileName(files(x).ToString.ToUpper)
                                                        xlWorkSheet.Cells(rowXls + 1, 13) = "DIGDAT\SFACILITY\SITUATIONMAP"
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = fileSize.ToString
                                                    End If

                                                    'xlWorkSheet.Cells(rowXls + 1, 20) = getApprovalGroupName(assetName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 27) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 28) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = 1

                                                ElseIf cbForm.SelectedValue.ToString = "SP_DIGITAL_DATA" Then

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = Path.GetFileName(files(x).ToString.ToUpper)
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = "DIGDAT\SFACILITY\BLOCKSTATION"
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 5) = getLoadBy(Login.TextBox1.Text)
                                                    'xlWorkSheet.Cells(rowXls + 1, 18) = getApprovalGroupName(assetName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 19) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = 1

                                                ElseIf cbForm.SelectedValue.ToString = "PROCESS_FACILITY_MAP" Then

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = Path.GetFileName(files(x).ToString.ToUpper)
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = "DIGDAT\SFACILITY\BLOCKSTATION"
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 4) = getTitle(Path.GetFileName(files(x).ToString.ToUpper), 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                    'xlWorkSheet.Cells(rowXls + 1, 13) = getApprovalGroupName(assetName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 22) = 1

                                                ElseIf cbForm.SelectedValue.ToString = "REGIONAL_MAP_IMAGE" Then

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = Path.GetFileName(files(x).ToString.ToUpper)
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = "DIGDAT\OTHERS"
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = getTitle(Path.GetFileName(files(x).ToString.ToUpper), 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 4) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 5) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 6) = getLoadBy(Login.TextBox1.Text)
                                                    'xlWorkSheet.Cells(rowXls + 1, 13) = getApprovalGroupName(assetName, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG", 1)
                                                    xlWorkSheet.Cells(rowXls + 1, 19) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = 1

                                                End If
                                                rowXls += 1
                                            End If
                                        End If
                                    Catch ex As Exception
                                        log(x + 1, "Error", Path.GetFileName(files(x).ToString) + ";" + ex.Message.ToString)
                                        CountErr += 1
                                    End Try
                                Next
                                ToolStripProgressBar((100 / files.Length) * (x + 1 - CountErr))
                                ToolStripStatusLabelTxt2("File " & x + 1 - CountErr & " of " & files.Length - CountErr)
                            Next
                            If Not IsNothing(xlApp) Then
                                Dim appDom As String = AppDomain.CurrentDomain.BaseDirectory + "output\"
                                xlApp.DisplayAlerts = False
                                Try
                                    xlWorkBook.SaveAs(appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + cbForm.SelectedValue.ToString + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                                xlWorkBook.Close(True, misValue, misValue)
                                xlApp.DisplayAlerts = True
                                xlApp.Quit()
                                Dim dateEnd As Date = Date.Now
                                End_Excel_App(dateStart, dateEnd)

                                releaseObject(xlWorkSheet)
                                releaseObject(xlWorkBook)
                                releaseObject(xlApp)

                                'Create listview
                                Dim filePath As String = appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + cbForm.SelectedValue.ToString + ".xls"

                                Try
                                    FillDataGriedView(filePath)
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End Try
                            End If

                            If CountErr > 0 Then
                                MessageBox.Show("Searching complete, " + files.Length.ToString + " Files found, With " + CountErr.ToString + " Error Found" + Environment.NewLine + "Please ensure that you fill the fields which highlight by yellow color!", "Finish", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Else
                                MessageBox.Show("Searching complete, " + files.Length.ToString + " Files found" + Environment.NewLine + "Please ensure that you fill the fields which highlight by yellow color!", "Finish", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If

                    End Select
                Else
                    MessageBox.Show("Directory does not exist", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Else
                MessageBox.Show("Directory cannot be null", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        ElseIf browseStatus = 2 Then
            respon = MessageBox.Show("Are you sure ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            Select Case respon
                Case vbYes
                    Try
                        FillDataGriedView(txBrowseFile.Text)
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
            End Select
        End If

        formType = cbForm.SelectedValue
        ToolStripStatusLabelTxt1("Ready")

        btnBrowse.Enabled = True
        btnBrowseFile.Enabled = True
        btnExecute.Enabled = True
        btnExport.Enabled = True
        btnUpload.Enabled = True
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub log(ByVal err_s As Integer, ByVal err_num As String, ByVal logMessage As String)
        If Not String.IsNullOrEmpty(filelogErr) Then
            Using sw As StreamWriter = File.AppendText(filelogErr)
                sw.Write("{0}; {1}", err_s, DateTime.Now.ToLongTimeString())
                sw.WriteLine(";{0};{1}", err_num, logMessage)
                sw.Flush()
                sw.Close()
            End Using
        End If
    End Sub

    Function getWell(ByVal strVal As String) As String
        Dim strHasil As String = ""
        fString = 0

        ' Ekstrak nama well sebelum underscore
        If strVal.Contains("_") Then
            strHasil = strVal.Substring(0, strVal.IndexOf("_"))
        End If

        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = "SELECT well_name FROM well WHERE well_name = :wellname"
            Dim ds As New DataSet()

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    cmd.Parameters.Add("wellname", OracleDbType.Varchar2).Value = strHasil

                    Using da As New OracleDataAdapter(cmd)
                        conn.Open()
                        da.Fill(ds)
                        conn.Close()
                    End Using
                End Using
            End Using

            If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                fString = 0 ' ditemukan
            Else
                fString = 2 ' tidak ditemukan
            End If

        Catch ex As Exception
            log(CountErr + 1, "DB Error", strVal + ";" + ex.Message)
            CountErr += 1
        End Try

        Return strHasil
    End Function


    Function getStructure(ByVal strVal As String, Optional ByVal flag As Integer = 0) As String
        ' Flag Code:
        ' 0 : cari aset_id berdasarkan s_struktur_alias
        ' 1 : cari structure_s berdasarkan s_struktur_alias
        ' 2 : cari structure_s berdasarkan nama_struktur
        ' 3 : cari aset_id berdasarkan nama_struktur

        Dim strHasil As String = String.Empty
        Dim find As String = String.Empty
        Dim eqWith As String = String.Empty
        fString = 0

        If strVal.Contains("_") Then
            strHasil = strVal.Substring(0, strVal.IndexOf("_"))
        End If

        Select Case flag
            Case 0
                find = "aset_id"
                eqWith = "s_struktur_alias"
            Case 1
                find = "structure_s"
                eqWith = "s_struktur_alias"
            Case 2
                find = "structure_s"
                eqWith = "nama_struktur"
            Case 3
                find = "aset_id"
                eqWith = "nama_struktur"
        End Select

        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = $"SELECT {find} FROM structure WHERE {eqWith} = :param"

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    cmd.Parameters.Add("param", OracleDbType.Varchar2).Value = strHasil

                    Dim adapter As New OracleDataAdapter(cmd)
                    Dim resultSet As New DataSet()

                    conn.Open()
                    adapter.Fill(resultSet)
                    conn.Close()

                    If resultSet.Tables.Count > 0 AndAlso resultSet.Tables(0).Rows.Count > 0 Then
                        strHasil = resultSet.Tables(0).Rows(0)(0).ToString()
                    End If
                End Using
            End Using

        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message)
            CountErr += 1
        End Try

        Return strHasil
    End Function


    Private Function getAsset(ByVal strVal As String) As String
        Dim strHasil As String = String.Empty
        fString = 0

        If strVal.Contains("_") Then
            Dim strSplit As String() = strVal.Split("_"c)
            If strSplit.Length > 2 Then
                strHasil = strSplit(2)
            End If
        End If

        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = "SELECT aset_id FROM aset WHERE aset_desc = :asetDesc"

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    cmd.Parameters.Add("asetDesc", OracleDbType.Varchar2).Value = strHasil

                    Dim adapter As New OracleDataAdapter(cmd)
                    Dim resultSet As New DataSet()

                    conn.Open()
                    adapter.Fill(resultSet)
                    conn.Close()

                    If resultSet.Tables.Count > 0 AndAlso resultSet.Tables(0).Rows.Count > 0 Then
                        strHasil = resultSet.Tables(0).Rows(0)("aset_id").ToString()
                    End If
                End Using
            End Using

        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message)
            CountErr += 1
        End Try

        Return strHasil
    End Function


    Private Function getRATitle(ByVal strVal As String) As String
        Dim strHasil As String = String.Empty
        fString = 0
        If (strVal.IndexOf("_") <> -1) Then
            Dim strSplit As String() = strVal.Split(New Char() {"_"})
            Dim dt As DateTime = DateTime.ParseExact(strSplit(3).Substring(0, strSplit(3).IndexOf(".")), "yyyyMMdd", CultureInfo.InvariantCulture)
            Dim dtReformat As String = dt.ToString("yyyy MMM", CultureInfo.InvariantCulture)
            strHasil = strSplit(0) + " " + strSplit(1) + " " + strSplit(2) + " " + dtReformat
        End If

        Return strHasil
    End Function

    Private Function getRA_Fdate(ByVal strVal As String) As String
        Dim strHasil As String = String.Empty
        fString = 0
        If (strVal.IndexOf("_") <> -1) Then
            Dim strSplit As String() = strVal.Split(New Char() {"_"})
            Dim dt As DateTime = DateTime.ParseExact(strSplit(3).Substring(0, strSplit(3).IndexOf(".")), "yyyyMMdd", CultureInfo.InvariantCulture)
            Dim dtReformat As String = dt.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)
            strHasil = dtReformat
        End If

        Return strHasil
    End Function

    Private Function getWellLogS() As String
        Dim strHasil As String = String.Empty
        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = "SELECT MAX(well_log_s) AS id FROM well_log"

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    conn.Open()
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot DBNull.Value Then
                        strHasil = result.ToString()
                    End If
                End Using
            End Using

        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message)
            CountErr += 1
        End Try

        Return strHasil
    End Function


    Private Function getGNGRepS() As String
        Dim strHasil As String = String.Empty
        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = "SELECT MAX(gng_report_s) AS id FROM gng_report"

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    conn.Open()
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot DBNull.Value Then
                        strHasil = result.ToString()
                    End If
                End Using
            End Using

        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message)
            CountErr += 1
        End Try

        Return strHasil
    End Function

    Private Function getSMI_S() As String
        Dim strHasil As String = String.Empty

        Try
            Dim jum_rec As Int16
            objDataSet.Reset()

            Dim sql_cek As String = "SELECT MAX(situation_map_image_s) AS id FROM situation_map_image"
            Using conn As New OracleConnection(connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql_cek, conn)
                    Using da As New OracleDataAdapter(cmd)
                        da.Fill(objDataSet, 0)
                    End Using
                End Using
            End Using

            jum_rec = objDataSet.Tables(0).Rows.Count
            If jum_rec > 0 Then
                strHasil = objDataSet.Tables(0).Rows(0)("id").ToString()
            End If

        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message.ToString())
            CountErr += 1
        End Try

        Return strHasil
    End Function

    Private Function getPFM_ID() As String
        Dim strHasil As String = String.Empty

        Try
            Dim jum_rec As Int16
            objDataSet.Reset()

            Dim sql_cek As String = "SELECT MAX(TO_NUMBER(pfm_drawing_id)) AS id FROM process_facility_map"
            Using conn As New OracleConnection(connectionString)
                conn.Open()
                Using cmd As New OracleCommand(sql_cek, conn)
                    Using da As New OracleDataAdapter(cmd)
                        da.Fill(objDataSet, 0)
                    End Using
                End Using
            End Using

            jum_rec = objDataSet.Tables(0).Rows.Count
            If jum_rec > 0 Then
                strHasil = objDataSet.Tables(0).Rows(0)("id").ToString()
            End If

        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message.ToString())
            CountErr += 1
        End Try

        Return strHasil
    End Function

    Function getWfType(ByVal strVal As String) As String
        Dim strHasil As String = String.Empty
        fString = 0

        strHasil = System.IO.Path.GetFileName(strVal).Replace("-", "_").Replace(" ", "")

        If strHasil.Contains("^") Then
            Dim strSpecial = strHasil.Substring(0, 1)
            If strHasil.Contains("_") Then
                strHasil = strHasil.Replace(strSpecial, "")
                Dim strSplit As String() = strHasil.Split("_"c)
                strHasil = strSplit(0)
            End If

            Try
                objDataSet.Reset()
                Dim sql_cek As String = $"SELECT r_document_type_id FROM elnusa.ref_report_type WHERE r_document_type_id = '{strHasil}'"
                Using conn As New OracleConnection(connectionString)
                    conn.Open()
                    Using cmd As New OracleCommand(sql_cek, conn)
                        Using da As New OracleDataAdapter(cmd)
                            da.Fill(objDataSet, "datacek")
                        End Using
                    End Using
                End Using

                If objDataSet.Tables("datacek").Rows.Count > 0 Then
                    fString = 0
                Else
                    fString = 1
                End If
            Catch ex As Exception
                log(CountErr + 1, "db error", strVal + ";" + ex.Message.ToString())
                CountErr += 1
            End Try
        Else
            strHasil = String.Empty
        End If

        Return strHasil
    End Function

    Function getWfSbj(ByVal strVal As String) As String
        Dim strHasil As String = String.Empty
        fString = 0

        strHasil = System.IO.Path.GetFileName(strVal).Replace("-", "_").Replace(" ", "")

        If strHasil.Contains("^") Then
            Dim strSpecial = strHasil.Substring(0, 1)
            If strHasil.Contains("_") Then
                strHasil = strHasil.Replace(strSpecial, "")
                Dim strSplit As String() = strHasil.Split("_"c)
                strHasil = If(strSplit.Length > 1, strSplit(1), "")
            End If

            Try
                objDataSet.Reset()
                Dim sql_cek As String = $"SELECT R_REPORT_SUBJECT_ID FROM elnusa.REF_REPORT_SUBJECT WHERE R_REPORT_SUBJECT_ID = '{strHasil}'"
                Using conn As New OracleConnection(connectionString)
                    conn.Open()
                    Using cmd As New OracleCommand(sql_cek, conn)
                        Using da As New OracleDataAdapter(cmd)
                            da.Fill(objDataSet, "dataCek")
                        End Using
                    End Using
                End Using

                If objDataSet.Tables("dataCek").Rows.Count > 0 Then
                    fString = 0
                Else
                    fString = 1
                End If
            Catch ex As Exception
                log(CountErr + 1, "DB Error", strVal + ";" + ex.Message.ToString())
                CountErr += 1
            End Try
        Else
            strHasil = String.Empty
        End If

        Return strHasil
    End Function


    Function getAuthor(ByVal strVal As String) As String
        Dim strHasil As String = String.Empty
        fString = 0

        Try
            ' Ambil nama file dari path
            Dim fileName As String = strVal.Substring(strVal.LastIndexOf("\") + 1)
            strHasil = fileName.Replace("-", "_").Replace(" ", "")

            If Not strHasil.Contains("^") Then
                Return String.Empty
            End If

            ' Ambil karakter pertama, lalu hilangkan dari string
            Dim specialChar As String = strHasil.Substring(0, 1)
            strHasil = strHasil.Replace(specialChar, "")

            ' Split berdasarkan "_"
            Dim strSplit As String() = strHasil.Split("_"c)
            If strSplit.Length = 3 Then
                Dim raUID As String = strSplit(2)

                ' Koneksi langsung pakai connection string Anda
                Dim connStr As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

                Using conn As New OracleConnection(connStr)
                    conn.Open()
                    Dim sql As String = "SELECT R_AUTHORS_NAME FROM elnusa.ref_authors WHERE RA_UID = :raUID"
                    Using cmd As New OracleCommand(sql, conn)
                        cmd.Parameters.Add(New OracleParameter("raUID", raUID))

                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                            Return result.ToString()
                        Else
                            fString = 1
                            Return String.Empty
                        End If
                    End Using
                End Using
            End If

        Catch ex As Exception
            log(CountErr + 1, "db error", strVal & ";" & ex.Message)
            CountErr += 1
        End Try

        Return String.Empty
    End Function

    Function getTitle(ByVal strVal As String, Optional ByVal flag As Integer = 0) As String
        Dim strHasil As String = ""
        fString = 0

        If flag = 0 Then
            Dim rex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("(WR|WRP|WRR|WL|WLR|WLP|IM|IML|IMR|TR)(([0-9][0-9][0-9][0-9][0-9])|([0-9][0-9][0-9][0-9][0-9][0-9]))(CD|OP|OT|T|A|E|F|G|H|L|M)")
            Dim matchCnt As Int16 = rex.Matches(strVal).Count
            Dim matchrx As System.Text.RegularExpressions.Match = rex.Match(strVal)
            If matchrx.Success And matchCnt = 1 Then
                strHasil = matchrx.Value
            ElseIf matchrx.Success And matchCnt > 1 Then
                fString = 2
                strHasil = matchrx.Value
            Else
                'Cari WRI|WLI
                rex = New System.Text.RegularExpressions.Regex("(WRI|WLI)")
                matchCnt = rex.Matches(strVal).Count
                matchrx = rex.Match(strVal)
                If matchrx.Success And matchCnt >= 1 Then
                    fString = 2
                    strHasil = matchrx.Value
                Else
                    'Cari Ext
                    rex = New System.Text.RegularExpressions.Regex("(.PDF|.TIF|.JPG)")
                    matchCnt = rex.Matches(strVal).Count
                    matchrx = rex.Match(strVal)
                    If matchrx.Success And matchCnt >= 1 Then
                        fString = 2
                        strHasil = matchrx.Value
                    Else
                        fString = 1
                        strHasil = matchrx.Value
                    End If
                End If
            End If
        End If

        If (strVal.IndexOf("_") <> -1) Then
            strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf("_" + strHasil) - strVal.IndexOf("_")).Replace("_", " ")
            fString = 0
        Else
            fString = 1
        End If
        Return strHasil
    End Function

    Function getCtLas(ByVal strVal As String) As String
        Dim strHasil As String = ""
        Dim strCek As String = ""
        fString = 0
cekLoop:
        Dim rex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("_([0-9])_")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As System.Text.RegularExpressions.Match = rex.Match(strVal)
        If matchCnt > 1 Then
            strCek = matchrx.Value
            strVal = strVal.Substring(0, strVal.LastIndexOf(strCek))
            GoTo cekLoop
        End If

        If matchrx.Success Then
            strCek = matchrx.Value
            strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf(strCek) - strVal.IndexOf("_")).Replace("_", " ").Trim
        Else
            rex = New System.Text.RegularExpressions.Regex("(_| )(([0-9])|([0-2][0-9])|([3][0-1]))\-(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\-\d{4}")
            matchrx = rex.Match(strVal)
            If matchrx.Success Then
                strCek = matchrx.Value
                strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf(strCek) - strVal.IndexOf("_")).Replace("_", " ").Trim
            Else
                rex = New System.Text.RegularExpressions.Regex("_WLD.LAS|_WLD \(")
                matchrx = rex.Match(strVal)
                If matchrx.Success Then
                    strCek = matchrx.Value
                    strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf(strCek) - strVal.IndexOf("_")).Replace("_", " ").Trim
                Else
                    fString = 1
                End If
            End If
        End If

        Return strHasil
    End Function

    Function getCtTif(ByVal strVal As String) As String
        Dim strHasil As String = ""
        Dim strCek As String = ""
        fString = 0
cekLoop:
        Dim rex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("_([0-9])-")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As System.Text.RegularExpressions.Match = rex.Match(strVal)
        If matchCnt > 1 Then
            strCek = matchrx.Value
            strVal = strVal.Substring(0, strVal.LastIndexOf(strCek))
            GoTo cekLoop
        End If

        If matchrx.Success Then
            strCek = matchrx.Value
            strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf(strCek) - strVal.IndexOf("_")).Replace("_", " ").Trim
        Else
            rex = New System.Text.RegularExpressions.Regex("(_| )(([0-9])|([0-2][0-9])|([3][0-1]))\-(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\-\d{4}")
            matchrx = rex.Match(strVal)
            If matchrx.Success Then
                strCek = matchrx.Value
                strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf(strCek) - strVal.IndexOf("_")).Replace("_", " ").Trim
            Else
                rex = New System.Text.RegularExpressions.Regex("_WLI.TIF|_WLI.PDF|_WLI.PDS|_WLI.JPG|_WLI \(")
                matchrx = rex.Match(strVal)
                If matchrx.Success Then
                    strCek = matchrx.Value
                    strHasil = strVal.Substring(strVal.IndexOf("_") + 1, strVal.LastIndexOf(strCek) - strVal.IndexOf("_")).Replace("_", " ").Trim
                Else
                    fString = 1
                End If
            End If
        End If

        Return strHasil
    End Function

    Function getDate(ByVal strVal As String) As String
        Dim strHasil As String = ""
        fString = 0
        Dim rex As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("(([0-9])|([0-2][0-9])|([3][0-1]))\-(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\-\d{4}")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As Text.RegularExpressions.Match = rex.Match(strVal)
        If matchrx.Success And matchCnt = 1 Then
            strHasil = matchrx.Value
        ElseIf matchrx.Success And matchCnt > 1 Then
            fString = 2
            strHasil = matchrx.Value
        Else
            fString = 1
        End If

        Return strHasil
    End Function

    Function getScale(ByVal strVal As String) As String
        Dim strHasil As String = ""
        fString = 0
        Dim rex As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("([0-9])\-(\d{4}|\d{3})")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As Text.RegularExpressions.Match = rex.Match(strVal)
        If matchrx.Success And matchCnt = 1 Then
            strHasil = "'" + matchrx.Value.Replace("-", ":")
        ElseIf matchrx.Success And matchCnt > 1 Then
            fString = 2
            strHasil = "'" + matchrx.Value.Replace("-", ":")
cekLoop:
            strVal = strVal.Replace(matchrx.Value, "")
            matchCnt = rex.Matches(strVal).Count
            matchrx = rex.Match(strVal)
            If matchrx.Success And matchCnt = 1 Then
                strHasil = strHasil + "; " + matchrx.Value.Replace("-", ":")
            ElseIf matchCnt > 1 Then
                GoTo cekLoop
            End If
        Else
            fString = 1
        End If
        Return strHasil
    End Function

    Function getRunNo(ByVal strVal As String) As String
        Dim strHasil As String = ""
        fString = 0
        Dim rex As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("_([0-9])_")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As Text.RegularExpressions.Match = rex.Match(strVal)
        If matchrx.Success And matchCnt = 1 Then
            strHasil = matchrx.Value.Replace("_", "")
        ElseIf matchrx.Success And matchCnt > 1 Then
            fString = 2
            strHasil = matchrx.Value.Replace("_", "")
        Else
            fString = 1
        End If

        Return strHasil
    End Function

    Function getTipe(ByVal strVal As String) As String
        Dim strHasil As String = ""
        fString = 0
        Dim rex As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("(_(WRI|WLI|HDR|SGI).|(IM|IML|IMR)(([0-9][0-9][0-9][0-9][0-9])|([0-9][0-9][0-9][0-9][0-9][0-9]))(OP))")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As Text.RegularExpressions.Match = rex.Match(strVal)
        If matchrx.Success And matchCnt = 1 Then
            If matchrx.Value = "_WRI." Then
                strHasil = "Well Report"
            ElseIf matchrx.Value = "_HDR." Then
                strHasil = "Well Log Header"
            ElseIf matchrx.Value = "_SGI." Then
                strHasil = "Technical Report"
            ElseIf matchrx.Value = "_WLI." Then
                strHasil = "Well Log Image"
            ElseIf matchrx.Value.StartsWith("IM") Then
                strHasil = "MAP"
            End If
        ElseIf matchrx.Success And matchCnt > 1 Then
            fString = 1
            If matchrx.Value = "_WRI." Then
                strHasil = "Well Report"
            ElseIf matchrx.Value = "_HDR." Then
                strHasil = "Well Log Header"
            ElseIf matchrx.Value = "_SGI." Then
                strHasil = "Technical Report"
            ElseIf matchrx.Value = "_WLI." Then
                strHasil = "Well Log Image"
            ElseIf matchrx.Value.StartsWith("IM") Then
                strHasil = "MAP"
            End If
        Else
            fString = 1
        End If

        Return strHasil
    End Function

    Public Function getNumberOfPdfPages(ByVal fileName As String) As Integer
        Try
            Dim matches As Text.RegularExpressions.MatchCollection
            Using sr As New StreamReader(File.OpenRead(fileName))
                Dim regex As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("/Type\s*/Page[^s]")
                matches = regex.Matches(sr.ReadToEnd())
            End Using

            Dim matchrx As Text.RegularExpressions.Match
            Using sr As New StreamReader(File.OpenRead(fileName))
                Dim regex2 As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("/N (\d{4}|\d{3}|\d{2}|\d{1})")
                matchrx = regex2.Match(sr.ReadToEnd())
            End Using

            Dim matchrx2 As Text.RegularExpressions.Match
            Using sr As New StreamReader(File.OpenRead(fileName))
                Dim regex3 As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("/Count (\d{4}|\d{3}|\d{2}|\d{1})/Type/Page")
                matchrx2 = regex3.Match(sr.ReadToEnd())
            End Using

            If matchrx2.Success Then
                Return matchrx2.Value.Replace("/Count ", "").Replace("/Type/Page", "")
            Else
                If matchrx.Success Then
                    If (matchrx.Value.Replace("/N ", "") = matches.Count.ToString) Then
                        Return matches.Count
                    Else
                        fString = 1
                        Return -1
                    End If
                Else
                    Return matches.Count
                End If
            End If

        Catch ex As Exception
            fString = 1
            Return -1
        End Try

    End Function

    Private Function pageCountPDF(ByRef pdfFile As String) As Integer
        ' Function for finding the number of pages in a given PDF file
        ' based on code found at http://www.dotnetspider.com/resources/21866-Count-pages-PDF-file.aspx

        pageCountPDF = 0

        Dim fileinfo = New FileInfo(pdfFile)
        If fileinfo.Exists Then
            Dim fs As FileStream = New FileStream(fileinfo.FullName, FileMode.Open, FileAccess.Read)
            Dim sr As StreamReader = New StreamReader(fs)
            Dim pdfMagicNumber() As Char = ("0000").ToCharArray

            sr.Read(pdfMagicNumber, 0, 4) ' put the first for characters of 
            ' the file into the pdfMagicNumber array

            If pdfMagicNumber = ("%PDF").ToCharArray Then 'The first four characters 
                ' of a PDF file should start with %PDF
                Dim pdfContents As String = sr.ReadToEnd()
                Dim rx As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("/Type\s/Page[^s]")
                Dim match As Text.RegularExpressions.MatchCollection = rx.Matches(pdfContents)
                pageCountPDF = match.Count
            Else
                Throw New Exception("File does not appear to be a PDF file (magic number not found).")
            End If
        Else
            Throw New Exception("File does not exist.")
        End If
    End Function

    Function getBarcode(ByVal strVal As String) As String
        Dim strHasil As String = ""
        fString = 0
        Dim rex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("(WR|WRP|WRR|WL|WLR|WLP|IM|IML|IMR|TR)(([0-9][0-9][0-9][0-9][0-9])|([0-9][0-9][0-9][0-9][0-9][0-9]))(CD|OP|OT|T|A|E|F|G|H|L|M)")
        Dim matchCnt As Int16 = rex.Matches(strVal).Count
        Dim matchrx As System.Text.RegularExpressions.Match = rex.Match(strVal)
        If matchrx.Success And matchCnt = 1 Then
            strHasil = matchrx.Value
        ElseIf matchrx.Success And matchCnt > 1 Then
            fString = 2
            strHasil = matchrx.Value
        Else
            'Cari lagi untuk barcode salah
            rex = New System.Text.RegularExpressions.Regex("((WR|WRP|WRR|WL|WLR|WLP|IM|IML|IMR|TR)(([0-9][0-9][0-9][0-9][0-9])|([0-9][0-9][0-9][0-9][0-9][0-9])))|((([0-9][0-9][0-9][0-9][0-9])|([0-9][0-9][0-9][0-9][0-9][0-9]))(CD|OP|OT|T|A|E|F|G|H|L|M))")
            matchCnt = rex.Matches(strVal).Count
            matchrx = rex.Match(strVal)
            If matchrx.Success And matchCnt >= 1 Then
                fString = 2
                strHasil = ""
            Else
                fString = 1
                strHasil = "NONE"
            End If
        End If

        Return strHasil
    End Function

    Function getCheck(ByVal strVal As String) As String
        Dim strHasil As String = strVal
        fString = 0
        Dim rex As Text.RegularExpressions.Regex = New Text.RegularExpressions.Regex("(([0-9])|([0-2][0-9])|([3][0-1]))-(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)(-)\d{4}")
        Dim matchrx As Text.RegularExpressions.Match = rex.Match(strVal)
        If matchrx.Success Then
            fString = 0
        Else
            fString = 1
            rex = New Text.RegularExpressions.Regex("(\&|\'|\%|\+|\#)")
            matchrx = rex.Match(strVal)
            If matchrx.Success Then
                fString = 2
            End If
        End If
        Return strHasil
    End Function

    Public Sub CreateFile()
        ''''modul tulis file > START
        If Len(bln) = 1 Then
            bln = "0" + bln
        End If
        If Len(tgll) = 1 Then
            tgll = "0" + tgll
        End If
        filelogErr = AppDomain.CurrentDomain.BaseDirectory + "output\" + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_Log_" + Replace(Replace(txBrowse.Text, ":\", "+"), "\", "-") + ".txt"

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
    End Sub

    Private Sub FindReplaceFillesToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        FindReplace.Show()
        Me.Hide()
    End Sub

    Private Sub RenameFilesToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        RenameFiles.ShowDialog()
    End Sub

    Private Sub GLV_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        Dim result As String = Login.ShowDialog()
        If result = DialogResult.OK Then
            If String.IsNullOrEmpty(My.Settings.Mode) Then
                My.Settings.Mode = 2
            End If

            If My.Settings.Mode = 1 Then
                Me.RadioButton1.Checked = True
                Me.Visible = True
                IsiCbForm()
                cbForm.Enabled = False
                _Del_BtnExecute(False)
            Else
                Me.RadioButton2.Checked = True
                Dim result2 As String = UNCLogin.ShowDialog()
                If result2 = DialogResult.OK Then
                    Me.Visible = True
                    IsiCbForm()
                    cbForm.Enabled = False
                    _Del_BtnExecute(False)
                Else
                    Close()
                End If
            End If
        Else
            Close()
        End If
    End Sub

    Private Sub IsiCbForm()
        Dim dict = New Dictionary(Of String, String)()
        dict.Add("[- - - - - - - -]", "[- - - - - - - -]")

        dict.Add("WELL_FILE", "WELL_FILE")
        dict.Add("WELL_LOG_DATA", "WELL_LOG_DATA")
        dict.Add("WELL_LOG_IMAGE", "WELL_LOG_IMAGE")
        dict.Add("WELL_MASTER_LOG", "WELL_MASTER_LOG")
        'dict.Add("WELL_CORRELATION", "WELL_CORRELATION")
        'dict.Add("GNG_CONTOUR_IMAGE", "GNG_CONTOUR_IMAGE")
        'dict.Add("GNG_REPORT", "GNG_REPORT")
        'dict.Add("REPO_WELL", "REPO_WELL")
        'dict.Add("REPO_ASET", "REPO_ASET")
        'dict.Add("SITUATION_MAP_IMAGE", "SITUATION_MAP_IMAGE")
        'dict.Add("PROCESS_FACILITY_MAP", "PROCESS_FACILITY_MAP")
        'dict.Add("SP_DIGITAL_DATA", "SP_DIGITAL_DATA")
        'dict.Add("REGIONAL_MAP_IMAGE", "REGIONAL_MAP_IMAGE")

        cbForm.DataSource = New BindingSource(dict, Nothing)
        cbForm.DisplayMember = "Value"
        cbForm.ValueMember = "Key"
    End Sub

    Private Sub cbForm_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cbForm.SelectedIndexChanged
        If cbForm.SelectedValue.ToString = "WELL_LOG_DATA" Then
            txType.Text = ".LAS"
        ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_IMAGE" Then
            txType.Text = ".TIF; .TIFF ; .PDS ; .PDF ; .JPG"
        ElseIf cbForm.SelectedValue.ToString = "WELL_FILE" Then
            txType.Text = ".PDF; .TIF ; .TIFF ; .JPG"
        ElseIf cbForm.SelectedValue.ToString = "WELL_MASTER_LOG" Then
            txType.Text = ".PDF; .TIF ; .TIFF ; .JPG ; .PNG"
        ElseIf cbForm.SelectedValue.ToString = "WELL_CORRELATION" Then
            txType.Text = ".PDF; .TIF ; .TIFF ; .JPG ; .PPT ; .PPTX"
        ElseIf cbForm.SelectedValue.ToString = "GNG_REPORT" Then
            txType.Text = ".PDF"
        ElseIf cbForm.SelectedValue.ToString = "GNG_CONTOUR_IMAGE" Then
            txType.Text = ".TIF; .TIFF; .PPT ; .PPTX"
        ElseIf cbForm.SelectedValue.ToString = "SITUATION_MAP_IMAGE" Then
            txType.Text = ".TIF; .TIFF; .PDF ; .JPG ; .DWG ; .ZIP"
        ElseIf cbForm.SelectedValue.ToString = "PROCESS_FACILITY_MAP" Then
            txType.Text = ".TIF; .TIFF; .DWG"
        ElseIf cbForm.SelectedValue.ToString = "SP_DIGITAL_DATA" Then
            txType.Text = ".TIF; .TIFF; .DWG"
        ElseIf cbForm.SelectedValue.ToString = "REPO_ASET" Then
            txType.Text = ".XLSX; .XLS"
        ElseIf cbForm.SelectedValue.ToString = "REGIONAL_MAP_IMAGE" Then
            txType.Text = ".PDF"
        End If
        btnExecute.Enabled = True
    End Sub

    Private Sub FillDataGriedView(ByVal filePath As String)
        _Del_BtnBrowse(False)
        _Del_BtnExecute(False)
        _Del_BtnRename(False)

        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()

        Dim xlApp As New Excel.Application
        'Dim xlApp As Object = CreateObject("Excel.Application")
        Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkBook As Object = xlApp.Workbooks.Add()
        Dim xlWorkSheet As Excel.Worksheet
        'Dim xlWorkSheet As Object = xlWorkBook.ActiveSheet
        Dim xlFile As String = filePath
        Dim fso As Object = CreateObject("Scripting.FileSystemObject")
        Dim oFile As Object
        Dim dt As DateTime
        Dim dateStart As Date = Date.Now
        'OldCulture = Thread.CurrentThread.CurrentCulture
        'newCulture = New CultureInfo(xlApp.LanguageSettings.LanguageID(MsoAppLanguageID.msoLanguageIDUI))
        'Thread.CurrentThread.CurrentCulture = newCulture
        oFile = fso.GetFile(xlFile)
        oFile.Attributes = 0
        'xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
        xlWorkBook = xlApp.Workbooks.Open(xlFile, , False)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'Thread.CurrentThread.CurrentCulture = OldCulture
        fString = 0

        'Design Datagridview Columns
        Dim xlTable As Excel.Range = xlWorkSheet.UsedRange
        DataGridView1.ColumnCount = xlTable.Columns.Count
        For colIndex As Integer = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(colIndex).Name = xlTable.Cells(1, colIndex + 1).Value
            ToolStripProgressBar(((ToolStripProgressBar1.Maximum / 2) / DataGridView1.Columns.Count) * (colIndex + 1))
            ToolStripStatusLabelTxt1("Processing..." & ToolStripProgressBar1.Value.ToString & "%")
        Next

        If cbForm.SelectedValue = "WELL_LOG_IMAGE" Then
            DataGridView1.Columns("WLI_FILE_NAME").Frozen = True
        ElseIf cbForm.SelectedValue = "WELL_LOG_DATA" Then
            DataGridView1.Columns("WLD_FILE_NAME").Frozen = True
        End If

        Dim tempProBarValue As Integer = ToolStripProgressBar1.Value

        'Populate DataGridView from Excel File
        DataGridView1.RowCount = xlTable.Rows.Count - 1
        'DataGridView1.Rows(rowIndex).HeaderCell.Value = rowIndex
        For colIndex As Integer = 0 To DataGridView1.Columns.Count - 1
            For rowIndex As Integer = 1 To xlTable.Rows.Count - 1

                If cbForm.SelectedValue.ToString = "WELL_FILE" Then
                    'Check WF_AUTHOR Column Value
                    If xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "WF_AUTHORS" Then
                        fString = 1
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbWFAuthor()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False

                        'Reformat WF_DATE 
                    ElseIf Not xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "WF_DATE" Then
                        dt = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = dt.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)

                        'Check WF_TYPE Column Value
                    ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "WF_TYPE" Then
                        fString = 2
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbWFType()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False

                        'Check WF_SUBJECT Column Value
                    ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "WF_SUBJECT" Then
                        fString = 3
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbWFSubject()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False

                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WF_NOTE" Then
                        fString = 3
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False

                        'Check Upload Flag
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If

                    Else
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                        'Yellow Color
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Interior.Color = 65535 Then
                            fString = 4
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow

                            'Orange Color
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Interior.Color = 42495 Then
                            fString = 5
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Orange
                        End If

                        'Duplicate Constraint Key
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Font.Color = 255 Then
                            fString = 6
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.ForeColor = Color.Red
                        End If

                        If xlTable._Default(rowIndex + 1, 7 + 1).Value = "PRE" And xlTable._Default(rowIndex + 1, 8 + 1).Value = "B3" Then
                            fString = 3
                            DataGridView1.Rows(rowIndex - 1).Cells("wf_file_path").Value = "DIGDAT\WELLTEST"
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                        End If

                        'Check Page Numer
                        Dim tes As String = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                        If xlTable.Cells(1, colIndex + 1).Value = "WF_NUM_OF_PAGE" Then
                            If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = "-1" Then
                                fString = 7
                            End If
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_DATA" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(1, colIndex + 1).Value = "WLD_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_REMARKS" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_PRODUCERS" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_TOP_DEPTH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_BOTTOM_DEPTH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_DEPTH_U" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_TITLE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_RUN_NO" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_LOG_TYPE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then
                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_IMAGE" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And (xlTable.Cells(1, colIndex + 1).Value = "WLI_FILE_NAME" Or xlTable.Cells(1, colIndex + 1).Value = "WLI_FILE_SIZE" Or xlTable.Cells(1, colIndex + 1).Value = "WLI_FILE_PATH" Or xlTable.Cells(1, colIndex + 1).Value = "WLI_LOAD_BY" Or xlTable.Cells(1, colIndex + 1).Value = "WLI_LOAD_DATE" Or xlTable.Cells(1, colIndex + 1).Value = "WELL_LOG_S" Or xlTable.Cells(1, colIndex + 1).Value = "WL_PRODUCERS" Or xlTable.Cells(1, colIndex + 1).Value = "WELL_NAME" Or xlTable.Cells(1, colIndex + 1).Value = "WELL_CONTRACTOR" Or xlTable.Cells(1, colIndex + 1).Value = "WL_TITLE") Then
                        fString = 1
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WLI_HDR_FILE_NAME" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WLI_HDR_FILE_PATH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WLI_HDR_FILE_SIZE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WLI_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WLI_VERTICAL_SCALE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_TOP_DEPTH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_BOTTOM_DEPTH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_REMARKS" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_DEPTH_U" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_TITLE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_RUN_NO" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WL_LOG_TYPE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then
                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "WELL_MASTER_LOG" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(1, colIndex + 1).Value = "WML_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WML_TOP_DEPTH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WML_BOTTOM_DEPTH" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WML_DEPTH_U" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WML_TITLE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WML_VERTICAL_SCALE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then
                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "WELL_CORRELATION" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(1, colIndex + 1).Value = "WC_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "WML_TITLE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "GNG_REPORT" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(1, colIndex + 1).Value = "GRI_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "GR_NUM_OF_PAGE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "GR_AUTHORS" Then
                        fString = 2
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbWFAuthor()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "GR_SUBJECT" Then
                        fString = 2
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbGRSubject()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "GNG_CONTOUR_IMAGE" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                ElseIf cbForm.SelectedValue.ToString = "REPO_WELL" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(1, colIndex + 1).Value = "RW_FTITLE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "RW_DESC" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "RW_FTYPE" Then
                        fString = 2
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbWFType()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "REPO_ASET" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(1, colIndex + 1).Value = "RA_FTITLE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "RA_DESC" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "RA_FSOURCE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "RA_FTYPE" Then
                        fString = 2
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbWFType()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "SITUATION_MAP_IMAGE" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "SMI_MAP_SUBJECT" Then
                        fString = 1
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbMapSubject()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "PROCESS_FACILITY_MAP" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "PFM_SUBJECT" Then
                        fString = 1
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbPFSubject()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "SP_DIGITAL_DATA" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "SP_ID" Then
                        fString = 1
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex) = cbSPID()
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False
                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                ElseIf cbForm.SelectedValue.ToString = "REGIONAL_MAP_IMAGE" Then
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                    DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                    If xlTable.Cells(rowIndex + 1, colIndex + 1).Value Is Nothing And xlTable.Cells(1, colIndex + 1).Value = "RMI_NOTE" Then
                        DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = False

                    ElseIf xlTable.Cells(1, colIndex + 1).Value = "FLAG" Then

                        'Default State
                        If xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 0 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Duplicate
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 1 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 2 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Insert Failed
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 3 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True

                            'Upload & Insert Success 
                        ElseIf xlTable.Cells(rowIndex + 1, colIndex + 1).Value = 4 Then
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = xlTable.Cells(rowIndex + 1, colIndex + 1).Value
                            DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                            DataGridView1.Rows(rowIndex - 1).Cells(colIndex).ReadOnly = True
                        End If
                    End If

                End If
            Next

            'Add number to row header
            For i = 0 To DataGridView1.RowCount - 1
                DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
            Next

            'Disable sort in column header
            For i = 0 To DataGridView1.Columns.Count - 1
                DataGridView1.Columns.Item(i).SortMode = DataGridViewColumnSortMode.Programmatic
            Next i

            ToolStripProgressBar(tempProBarValue + (((ToolStripProgressBar1.Maximum / 2) / DataGridView1.Columns.Count) * (colIndex + 1)))
            ToolStripStatusLabelTxt1("Processing..." & ToolStripProgressBar1.Value.ToString & "%")
        Next

        ToolStripStatusLabelTxt1("Complete")

        xlApp.Quit()
        Dim dateEnd As Date = Date.Now
        End_Excel_App(dateStart, dateEnd)

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        ToolStripStatusLabelTxt1("Ready")
        ToolStripProgressBar(0)

        If fString = 0 Then
            _Del_BtnExport(True)

        Else
            _Del_BtnExport(False)
        End If

        _Del_BtnBrowse(True)
        _Del_BtnExecute(True)
        _Del_BtnRename(True)
    End Sub

    Private Function cbWFAuthor()
        Dim cmb As New DataGridViewComboBoxCell
        query = "select r_authors_s, r_authors_name from ref_authors order by r_authors_name"
        Dim ds As DataSet = New DataSet()
        ds.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    conn.Open()
                    dataAdp = New OracleDataAdapter(comm)
                    dataAdp.Fill(ds, 0)
                    With cmb
                        .DataSource = ds.Tables(0)
                        .DisplayMember = "r_authors_name"
                        .ValueMember = "r_authors_name"
                        .FlatStyle = FlatStyle.Flat
                    End With
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return cmb
    End Function

    Private Function cbWFType()
        Dim cmb As New DataGridViewComboBoxCell
        query = "select r_document_type_id, r_document_type_id||'-'||r_document_type as r_document_type from ref_report_type order by r_document_type"
        Dim ds As DataSet = New DataSet()
        ds.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    conn.Open()
                    dataAdp = New OracleDataAdapter(comm)
                    dataAdp.Fill(ds, 0)
                    With cmb
                        .DataSource = ds.Tables(0)
                        .DisplayMember = "r_document_type"
                        .ValueMember = "r_document_type_id"
                        .FlatStyle = FlatStyle.Flat
                    End With
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return cmb
    End Function

    Private Function cbWFSubject()
        Dim cmb As New DataGridViewComboBoxCell
        query = "select r_report_subject_id, r_report_subject_id||'-'||r_report_subject_nm as r_report_subject_nm from ref_report_subject order by r_report_subject_id"
        Dim ds As DataSet = New DataSet()
        ds.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    conn.Open()
                    dataAdp = New OracleDataAdapter(comm)
                    dataAdp.Fill(ds, 0)
                    With cmb
                        .DataSource = ds.Tables(0)
                        .DisplayMember = "r_report_subject_nm"
                        .ValueMember = "r_report_subject_id"
                        .AutoComplete = True
                        .FlatStyle = FlatStyle.Flat
                    End With
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return cmb
    End Function

    Private Function cbMapSubject()
        Dim cmb As New DataGridViewComboBoxCell
        query = "select r_map_subject_id, r_map_subject_nm from ref_map_subject order by r_map_subject_id"
        Dim ds As DataSet = New DataSet()
        ds.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    conn.Open()
                    dataAdp = New OracleDataAdapter(comm)
                    dataAdp.Fill(ds, 0)
                    With cmb
                        .DataSource = ds.Tables(0)
                        .DisplayMember = "r_map_subject_nm"
                        .ValueMember = "r_map_subject_id"
                        .AutoComplete = True
                        .FlatStyle = FlatStyle.Flat
                    End With
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return cmb
    End Function

    Private Function cbPFSubject()
        Dim cmb As New DataGridViewComboBoxCell
        query = "select r_pf_subject_id, r_pf_subject_nm from ref_pf_subject order by r_pf_subject_id"
        Dim ds As DataSet = New DataSet()
        ds.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    conn.Open()
                    dataAdp = New OracleDataAdapter(comm)
                    dataAdp.Fill(ds, 0)
                    With cmb
                        .DataSource = ds.Tables(0)
                        .DisplayMember = "r_pf_subject_nm"
                        .ValueMember = "r_pf_subject_id"
                        .AutoComplete = True
                        .FlatStyle = FlatStyle.Flat
                    End With
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return cmb
    End Function

    Private Function cbSPID()
        Dim cmb As New DataGridViewComboBoxCell
        query = "select sp_id, sp_name from stasiun_pengumpul order by sp_id"
        Dim ds As DataSet = New DataSet()
        ds.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                End With
                Try
                    conn.Open()
                    dataAdp = New OracleDataAdapter(comm)
                    dataAdp.Fill(ds, 0)
                    With cmb
                        .DataSource = ds.Tables(0)
                        .DisplayMember = "sp_name"
                        .ValueMember = "sp_id"
                        .AutoComplete = True
                        .FlatStyle = FlatStyle.Flat
                    End With
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return cmb
    End Function

    Private Function cbGRSubject()
        Dim cmb As New DataGridViewComboBoxCell
        cmb.Items.Add("GGR")
        cmb.Items.Add("OTH")

        Return cmb
    End Function

    Private Sub ExportToXL(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim respon As String = MessageBox.Show("Are you sure you want to save?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        btnBrowse.Enabled = False
        btnExecute.Enabled = False
        btnExport.Enabled = False
        btnRename.Enabled = False
        btnUpload.Enabled = False

        fString = 0
        Dim misValue As Object = Reflection.Missing.Value

        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add(misValue)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets("sheet1")
        Dim appDom As String = AppDomain.CurrentDomain.BaseDirectory + "output\"

        Select Case respon
            Case vbYes
                'Export Header Names Start
                Dim columnsCount As Integer = DataGridView1.Columns.Count

                Dim j As Integer = 0
                For Each column In DataGridView1.Columns
                    xlWorkSheet.Cells(1, column.Index + 1).Value = column.Name
                    ToolStripProgressBar(((ToolStripProgressBar1.Maximum / 2) / DataGridView1.Columns.Count) * (j + 1))
                    ToolStripStatusLabelTxt1("Processing..." & ToolStripProgressBar1.Value.ToString & "%")
                    j += 1
                Next
                'Export Header Name End

                Dim tempProBarValue As Integer = ToolStripProgressBar1.Value

                'Export Each Row Start
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    Dim columnIndex As Integer = 0
                    Do Until columnIndex = columnsCount
                        If DataGridView1.Columns(columnIndex).Name = "WLI_VERTICAL_SCALE" And Not String.IsNullOrEmpty(DataGridView1.Item(columnIndex, i).Value) Then
                            xlWorkSheet.Cells(i + 2, columnIndex + 1).Value = "'" + DataGridView1.Item(columnIndex, i).Value.ToString
                        ElseIf DataGridView1.Columns(columnIndex).Name = "WML_VERTICAL_SCALE" And Not String.IsNullOrEmpty(DataGridView1.Item(columnIndex, i).Value) Then
                            xlWorkSheet.Cells(i + 2, columnIndex + 1).Value = "'" + DataGridView1.Item(columnIndex, i).Value.ToString
                        Else
                            xlWorkSheet.Cells(i + 2, columnIndex + 1).Value = DataGridView1.Item(columnIndex, i).Value
                        End If

                        'Check Flag Upload
                        If cbFormSelValue = "WELL_FILE" Then
                            xlWorkSheet.Cells(i + 2, 30) = 0
                            If Not String.IsNullOrEmpty(xlWorkSheet.Cells(i + 2, 4).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(i + 2, 5).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(i + 2, 6).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(i + 2, 8).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(i + 2, 9).Value) And xlWorkSheet.Cells(i + 2, 29).value = 0 Then
                                xlWorkSheet.Cells(i + 2, 30) = 1
                            End If
                        ElseIf cbFormSelValue = "REPO_ASET" Then
                            If Not String.IsNullOrEmpty(xlWorkSheet.Cells(i + 2, 2).Value) Then
                                xlWorkSheet.Cells(i + 2, 24) = 1
                            End If
                        End If

                        ToolStripProgressBar(tempProBarValue + (((ToolStripProgressBar1.Maximum / 2) / columnsCount) * (columnIndex + 1)))
                        ToolStripStatusLabelTxt1("Processing..." & ToolStripProgressBar1.Value.ToString & "%")
                        columnIndex += 1
                    Loop
                Next
                'Export Each Row End

                ToolStripStatusLabelTxt1("Complete")

                xlApp.DisplayAlerts = False
                Try
                    xlWorkBook.SaveAs(appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + cbForm.SelectedValue.ToString + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                xlWorkBook.Close(True, misValue, misValue)
                xlApp.DisplayAlerts = True
                xlApp.Quit()

                releaseObject(xlWorkSheet)
                releaseObject(xlWorkBook)
                releaseObject(xlApp)

                'Clear DataGridView
                DataGridView1.Rows.Clear()
                DataGridView1.ClearSelection()

                ToolStripProgressBar(0)
                ToolStripStatusLabelTxt1("Ready")

                'Load Data Again to DataGridView
                Dim filePath As String = appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + cbForm.SelectedValue.ToString + ".xls"
                FillDataGriedView(filePath)
                MessageBox.Show("The file has been saved succesfully", "Sucess", MessageBoxButtons.OK, MessageBoxIcon.Information)

                btnExecute.Enabled = True
                btnExport.Enabled = True
                btnBrowse.Enabled = True
                btnRename.Enabled = True
                btnUpload.Enabled = True

            Case vbNo
                btnExecute.Enabled = True
                btnExport.Enabled = True
                btnBrowse.Enabled = True
                btnRename.Enabled = True
                btnUpload.Enabled = True
        End Select
    End Sub

    Private Sub ExportToXLRep(uploadType As Integer)
        _Del_BtnBrowse(False)
        _Del_BtnExecute(False)
        _Del_BtnRename(False)
        _Del_BtnUpload(False)

        Dim misValue As Object = Reflection.Missing.Value

        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add(misValue)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets("sheet1")
        Dim appDom As String = AppDomain.CurrentDomain.BaseDirectory + "output\"
        Dim dateStart As Date = Date.Now

        Try
            'Export Header Names Start
            Dim columnsCount As Integer = DataGridView1.Columns.Count
            For Each column In DataGridView1.Columns
                xlWorkSheet.Cells(1, column.Index + 1).Value = column.Name
                ToolStripProgressBar(((ToolStripProgressBar1.Maximum / 2) / DataGridView1.Columns.Count) * (j + 1))
                ToolStripStatusLabelTxt1("Processing..." & ToolStripProgressBar1.Value.ToString & "%")
            Next
            'Export Header Name End

            Dim tempProBarValue As Integer = ToolStripProgressBar1.Value

            'Export Each Row Start
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                Dim columnIndex As Integer = 0
                Do Until columnIndex = columnsCount
                    If DataGridView1.Columns(columnIndex).Name = "WLI_VERTICAL_SCALE" And Not String.IsNullOrEmpty(DataGridView1.Item(columnIndex, i).Value) Then
                        xlWorkSheet.Cells(i + 2, columnIndex + 1).Value = "'" + DataGridView1.Item(columnIndex, i).Value.ToString
                    ElseIf DataGridView1.Columns(columnIndex).Name = "WML_VERTICAL_SCALE" And Not String.IsNullOrEmpty(DataGridView1.Item(columnIndex, i).Value) Then
                        xlWorkSheet.Cells(i + 2, columnIndex + 1).Value = "'" + DataGridView1.Item(columnIndex, i).Value.ToString
                    Else
                        xlWorkSheet.Cells(i + 2, columnIndex + 1).Value = DataGridView1.Item(columnIndex, i).Value
                    End If

                    ToolStripProgressBar(tempProBarValue + (((ToolStripProgressBar1.Maximum / 2) / columnsCount) * (columnIndex + 1)))
                    ToolStripStatusLabelTxt1("Processing..." & ToolStripProgressBar1.Value.ToString & " %")
                    columnIndex += 1
                Loop
            Next
            'Export Each Row End
            ToolStripStatusLabelTxt1("Complete")

            xlApp.DisplayAlerts = False
            Try
                xlWorkBook.SaveAs(appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + formType.ToString + "_REP.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.DisplayAlerts = True
            xlApp.Quit()

            Dim dateEnd As Date = Date.Now
            End_Excel_App(dateStart, dateEnd)

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)

            'Clear DataGridView
            DataGridView1.ClearSelection()

            ToolStripStatusLabelTxt1("Ready")
            ToolStripProgressBar(0)

            'Load Data Again to DataGridView
            'If rowIndx = rowCount Then
            Dim filePath As String = appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + formType.ToString + "_REP.xls"
            FillDataGriedView(filePath)
            'End If

            _Del_BtnBrowse(True)
            _Del_BtnExecute(True)
            _Del_BtnRename(True)
            _Del_BtnUpload(True)

        Catch e As Exception
            Dim err As String = e.Message
            MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.SelectedRows.Count > 0 Then
            'DataGridView1.ClearSelection()
            btnRename.Enabled = False
            btnOpenFile.Enabled = False
        Else
            If sender.columns(e.ColumnIndex).headercell.value = "WF_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WLD_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WLI_FILE_NAME" Or sender.columns(e.ColumnIndex).headercell.value = "WLI_HDR_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WML_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WC_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "GCI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "GRI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RW_FNAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RA_FNAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "SMI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "SPDD_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "PFM_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RMI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
            Else
                btnRename.Enabled = False
                btnOpenFile.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnRename_Click(sender As Object, e As EventArgs) Handles btnRename.Click
        Dim result As String = RenameFile.ShowDialog()
        If result = DialogResult.OK Then
            'DataGridView1.CurrentCell = activeCell
            'DataGridView1.BeginEdit(True)
        End If
    End Sub

    Private Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click
        Dim rowIndex As Integer = DataGridView1.CurrentCell.RowIndex
        Dim fullPath As String = DataGridView1.Rows(rowIndex).Cells("path_sourcefile").Value + "\" + DataGridView1.CurrentCell.Value


        If File.Exists(fullPath) = True Then
            Dim process As Process = New Process()
            Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
            startInfo.FileName = fullPath
            process.StartInfo = startInfo
            process.Start()
            'process.WaitForExit()
        Else
            MessageBox.Show("File does Not exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Function CheckPKWellFile(ByVal WFN As String, ByVal WLN As String, ByVal WFT As String, ByVal WFS As String) As String
        fString = 0
        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = "SELECT well_name FROM elnusa.well_file WHERE wf_file_name = :WFN AND well_name = :WLN AND wf_type = :WFT AND wf_subject = :WFS"
            Dim ds As New DataSet()

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    ' Tambahkan parameter untuk mencegah SQL Injection
                    cmd.Parameters.Add("WFN", OracleDbType.Varchar2).Value = WFN
                    cmd.Parameters.Add("WLN", OracleDbType.Varchar2).Value = WLN
                    cmd.Parameters.Add("WFT", OracleDbType.Varchar2).Value = WFT
                    cmd.Parameters.Add("WFS", OracleDbType.Varchar2).Value = WFS

                    Using da As New OracleDataAdapter(cmd)
                        conn.Open()
                        da.Fill(ds)
                        conn.Close()
                    End Using
                End Using
            End Using

            If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                fString = 1 ' Data ditemukan
            Else
                fString = 0 ' Tidak ditemukan
            End If

        Catch ex As Exception
            log(CountErr + 1, "DB Error :", ex.Message)
            CountErr += 1
        End Try

        Return fString
    End Function


    Private Function CheckPKWellLogImage(ByVal WFN As String) As String
        Dim result As String = "0"
        Dim connectionString As String = "User Id=elnusa;Password=elnusa;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.203.1.231)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=inametadb)))"

        Try
            Dim query As String = "SELECT wli_file_name FROM elnusa.well_log_image WHERE wli_file_name = :wfn"
            Dim ds As New DataSet()

            Using conn As New OracleConnection(connectionString)
                Using cmd As New OracleCommand(query, conn)
                    cmd.Parameters.Add(New OracleParameter("wfn", OracleDbType.Varchar2)).Value = WFN

                    Using da As New OracleDataAdapter(cmd)
                        conn.Open()
                        da.Fill(ds)
                        conn.Close()
                    End Using
                End Using
            End Using

            If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                result = "1"
            End If
        Catch ex As Exception
            log(CountErr + 1, "DB Error : ", ex.Message.ToString())
            CountErr += 1
        End Try

        Return result
    End Function


    Private Function getApprovalGroupName(ByVal searchVal As String, ByVal filename As String, ByVal dataGroup As String, Optional ByVal flag As Integer = 0) As String
        Dim strHasil As String = String.Empty
        Dim jum_rec As Int16
        Dim find As String = String.Empty
        If flag = 0 Then
            find = " LOOKUP.GET_ASET_ID(LOOKUP.GET_STRUCTURE_S(:find)) "
        Else
            find = " :find "
        End If
        query = "select user_id from REF_APPROVAL_GROUP where asset_id = " + find + " and data_group = :dataGroup"
        objDataSet.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.Add(":find", OracleDbType.Varchar2).Value = searchVal
                    .Parameters.Add(":dataGroup", OracleDbType.Varchar2).Value = dataGroup
                End With
                Try
                    dataAdp = New OracleDataAdapter(comm)
                    conn.Open()
                    dataAdp.Fill(objDataSet, 0)
                    jum_rec = objDataSet.Tables(0).Rows.Count
                    If jum_rec > 0 Then
                        strHasil = objDataSet.Tables(0).Rows(0)("user_id").ToString
                    End If
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return strHasil
    End Function

    Private Function getLoadBy(ByVal username As String) As String
        Dim strHasil As String = String.Empty
        Dim jum_rec As Int16
        query = "select usrinfo_initial from user_info where user_id = :username"
        objDataSet.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.Add(":username", OracleDbType.Varchar2).Value = username
                End With
                Try
                    dataAdp = New OracleDataAdapter(comm)
                    conn.Open()
                    dataAdp.Fill(objDataSet, 0)
                    jum_rec = objDataSet.Tables(0).Rows.Count
                    If jum_rec > 0 Then
                        strHasil = objDataSet.Tables(0).Rows(0)("usrinfo_initial").ToString
                    End If
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return strHasil
    End Function

    Private Function getElevetionUnit(ByVal wellname As String) As String
        Dim strHasil As String = String.Empty
        Dim jum_rec As Int16
        query = "select w_elevation_u from well where well_name = :wellname"
        objDataSet.Reset()

        Using conn As New OracleConnection(conString)
            Using comm As New OracleCommand()
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Parameters.Add(":wellname", OracleDbType.Varchar2).Value = wellname
                End With
                Try
                    dataAdp = New OracleDataAdapter(comm)
                    conn.Open()
                    dataAdp.Fill(objDataSet, 0)
                    jum_rec = objDataSet.Tables(0).Rows.Count
                    If jum_rec > 0 Then
                        strHasil = objDataSet.Tables(0).Rows(0)("w_elevation_u").ToString
                    End If
                    conn.Close()
                Catch ex As Exception
                    conn.Close()
                    MessageBox.Show(ex.Message)
                    Throw
                End Try
            End Using
        End Using

        Return strHasil
    End Function

    Private Sub CopyFiles()
        ToolStripProgressBar1.Minimum = 0
        ToolStripProgressBar1.Maximum = 100
        ToolStripProgressBar1.Value = 0
        If Not bgWorker.IsBusy Then
            btnUpload.Enabled = False
            bgWorker.RunWorkerAsync()
        Else
            MessageBox.Show("Application is busy, please wait a moment", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        btnCancel.Enabled = False
        _Del_BtnBrowse(True)
        _Del_BtnExecute(True)
        _Del_BtnExport(True)
        _Del_BtnUpload(True)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        _Del_BtnBrowse(False)
        _Del_BtnExecute(False)
        _Del_BtnExport(False)
        _Del_BtnUpload(False)

        respon = MessageBox.Show("Do you really want to upload the file(s) ?", "Confirmation sultan", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If respon = DialogResult.Yes Then
            CopyFiles()
        Else
            MessageBox.Show("Upload dibatalkan oleh pengguna.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            _Del_BtnUpload(True) ' kembalikan tombol jika tidak jadi upload
        End If
        _Del_BtnUpload(True)
    End Sub

    Private Sub bgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgWorker.DoWork
        Dim totalRows As Integer = DataGridView1.Rows.Cast(Of DataGridViewRow)().Count(Function(r) Not r.IsNewRow)
        Dim currentRow As Integer = 0

        ' Set Maximum ProgressBar di UI thread
        Me.Invoke(Sub()
                      ToolStripProgressBar1.Minimum = 0
                      ToolStripProgressBar1.Maximum = totalRows
                      ToolStripProgressBar1.Value = 0

                      ' Tambahkan kolom log_status jika belum ada
                      If Not DataGridView1.Columns.Contains("log_status") Then
                          DataGridView1.Columns.Add("log_status", "Log Status")
                      End If
                  End Sub)

        Using conn As New OracleConnection(connectionString)
            conn.Open()

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.IsNewRow Then Continue For

                currentRow += 1

                Dim pathSource As String = row.Cells("path_sourcefile").Value.ToString().Trim()
                Dim fileName As String = ""
                Dim fileName2 As String = ""
                Dim relativeTargetPath As String = ""
                Dim relativeTargetPath2 As String = ""

                Select Case cbFormSelValue
                    Case "WELL_FILE"
                        fileName = row.Cells("wf_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("wf_file_path").Value.ToString().Trim()
                    Case "WELL_LOG_DATA"
                        fileName = row.Cells("wld_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("wld_file_path").Value.ToString().Trim()
                    Case "WELL_LOG_IMAGE"
                        fileName = row.Cells("wli_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("wli_file_path").Value.ToString().Trim()
                    Case "WELL_MASTER_LOG"
                        fileName = row.Cells("wml_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("wml_file_path").Value.ToString().Trim()
                    Case "WELL_CORRELATION"
                        fileName = row.Cells("wc_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("wc_file_path").Value.ToString().Trim()
                    Case "GNG_CONTOUR_IMAGE"
                        fileName = row.Cells("gci_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("gci_file_path").Value.ToString().Trim()
                    Case "GNG_REPORT"
                        fileName = row.Cells("gri_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("gri_file_path").Value.ToString().Trim()
                    Case "REPO_WELL"
                        fileName = row.Cells("rw_fname").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("rw_fpath").Value.ToString().Trim()
                    Case "REPO_ASET"
                        fileName = row.Cells("ra_fname").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("ra_fpath").Value.ToString().Trim()
                    Case "SITUATION_MAP_IMAGE"
                        fileName = row.Cells("smi_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("smi_file_path").Value.ToString().Trim()
                    Case "SP_DIGITAL_DATA"
                        fileName = row.Cells("spdd_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("spdd_file_path").Value.ToString().Trim()
                    Case "PROCESS_FACILITY_MAP"
                        fileName = row.Cells("pfm_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("pfm_file_path").Value.ToString().Trim()
                    Case "REGIONAL_MAP_IMAGE"
                        fileName = row.Cells("rmi_file_name").Value.ToString().Trim()
                        relativeTargetPath = row.Cells("rmi_file_path").Value.ToString().Trim()
                    Case Else
                        Continue For ' Skip baris jika Form belum dikenali
                End Select

                ' Bangun target folder
                Dim targetFolder As String = My.Settings.DigdatHost & relativeTargetPath
                Dim targetFolder2 As String = My.Settings.DigdatHost & relativeTargetPath2
                If Not Directory.Exists(targetFolder) Then
                    Directory.CreateDirectory(targetFolder)
                End If

                Dim fullSourcePath2 As String = Path.Combine(pathSource, fileName2)
                Dim targetPath2 As String = Path.Combine(targetFolder2, fileName2)
                Dim fullSourcePath As String = Path.Combine(pathSource, fileName)
                Dim targetPath As String = Path.Combine(targetFolder, fileName)

                ' Proses copy file
                Try
                    If File.Exists(fullSourcePath) Then
                        If File.Exists(targetPath) Then
                            If Not replaceAll AndAlso Not skipAll Then
                                Dim result As DialogResult = DialogResult.None
                                Me.Invoke(Sub()
                                              result = MessageBox.Show($"File '{fileName}' sudah ada di tujuan.{vbCrLf}Ganti?", "Konfirmasi", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                                          End Sub)

                                If result = DialogResult.Cancel Then Exit For
                                If result = DialogResult.No Then Continue For

                                Dim askAll As DialogResult = DialogResult.None
                                Me.Invoke(Sub()
                                              askAll = MessageBox.Show("Terapkan pilihan ini untuk semua file berikutnya?", "Konfirmasi Semua", MessageBoxButtons.YesNo)
                                          End Sub)

                                If askAll = DialogResult.Yes Then
                                    If result = DialogResult.Yes Then replaceAll = True
                                    If result = DialogResult.No Then skipAll = True
                                End If
                            ElseIf skipAll Then
                                Continue For
                            End If
                        End If
                        Try
                            'File.Copy(fullSourcePath, targetPath, True)
                            Dim success As Boolean = CopyFileWithCancel(fullSourcePath, targetPath, bgWorker)
                            If cbFormSelValue = "WELL_LOG_IMAGE" Then
                                If row.Cells("wli_hdr_file_name").Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(row.Cells("wli_hdr_file_name").Value.ToString().Trim()) Then
                                    fileName2 = row.Cells("wli_hdr_file_name").Value.ToString().Trim()
                                    relativeTargetPath2 = row.Cells("wli_hdr_file_path").Value.ToString().Trim()
                                    'File.Copy(fullSourcePath2, targetPath2, True)
                                    Dim success2 As Boolean = CopyFileWithCancel(fullSourcePath2, targetPath2, bgWorker)
                                End If
                            End If

                            InsertToDatabase(row, conn)
                            successFlg = 1
                        Catch ex As Exception
                            Me.Invoke(Sub()
                                          MessageBox.Show("Insert DB gagal untuk file: " & fileName & vbCrLf & ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                      End Sub)
                        End Try
                    Else
                        Me.Invoke(Sub()
                                      MessageBox.Show("File tidak ditemukan: " & fullSourcePath, "File Hilang", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                  End Sub)
                    End If

                Catch ex As Exception
                    Me.Invoke(Sub()
                                  MessageBox.Show("Gagal menyalin file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              End Sub)
                End Try

                ' Update ProgressBar
                'Dim progress As Integer = currentRow
                'bgWorker.ReportProgress(progress, $"Menyalin: {fileName}")
                ' Update ProgressBar
                Dim progress As Integer = currentRow
                If progress > totalRows Then progress = totalRows
                If progress < 0 Then progress = 0
                bgWorker.ReportProgress(progress, $"Menyalin: {fileName}")

            Next
            conn.Close()
        End Using
    End Sub
    Private Sub bgWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles bgWorker.ProgressChanged
        'ToolStripProgressBar1.Value = Math.Min(ToolStripProgressBar1.Maximum, e.ProgressPercentage)
        'ToolStripStatusLabel1.Text = e.UserState.ToString() & " | " & e.ProgressPercentage.ToString() & "%"
        ToolStripProgressBar1.Value = e.ProgressPercentage
        ToolStripStatusLabel1.Text = e.UserState.ToString()
    End Sub
    Private Sub bgWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgWorker.RunWorkerCompleted
        btnUpload.Enabled = True
        If successFlg = 1 Then
            ToolStripStatusLabel1.Text = "Upload Success"
            ToolStripProgressBar1.Value = 0
            MessageBox.Show("Proses copy file selesai.", "Selesai", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ToolStripStatusLabel1.Text = "Upload Cancelled"
            ToolStripProgressBar1.Value = 0
            MessageBox.Show("Upload has been cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        _Del_BtnBrowse(True)
        _Del_BtnExecute(True)
        _Del_BtnExport(True)
        _Del_BtnUpload(True)
        _Del_BtnCancel(False)
    End Sub

    Private Function CopyFileWithCancel(sourceFile As String, targetFile As String, worker As BackgroundWorker) As Boolean
        Const bufferSize As Integer = 81920 ' 80 KB
        Dim buffer(bufferSize - 1) As Byte
        Dim bytesRead As Integer

        Try
            Using sourceStream As New FileStream(sourceFile, FileMode.Open, FileAccess.Read)
                Using targetStream As New FileStream(targetFile, FileMode.Create, FileAccess.Write)
                    Do
                        If worker.CancellationPending Then
                            Return False
                        End If

                        bytesRead = sourceStream.Read(buffer, 0, buffer.Length)
                        If bytesRead = 0 Then Exit Do

                        targetStream.Write(buffer, 0, bytesRead)
                    Loop
                End Using
            End Using

            Return True
        Catch ex As Exception
            ' Tambahkan log atau MessageBox jika perlu
            Return False
        End Try
    End Function
    Private Sub InsertToDatabase(row As DataGridViewRow, conn As OracleConnection)
        Try
            Select Case cbFormSelValue
                Case "WELL_FILE"
                    'Console.WriteLine(cbFormSelValue, "makan bang")
                    'MessageBox.Show("makan bang: ")
                    Dim query As String = "INSERT INTO well_file (" &
                    "wf_file_name, well_name, well_contractor, wf_barcode, wf_title, " &
                    "wf_authors, wf_date, wf_type, wf_subject, wf_group, " &
                    "wf_num_of_page, wf_note, wf_doc_ver, wf_file_path, wf_file_size, " &
                    "wf_load_by, wf_loaded_date, wf_verified_by, wf_verified_date, approval_group, " &
                    "approval_status, approval_inameta_by, approval_inameta_date, approval_inameta_note, " &
                    "approval_user_by, approval_user_date, approval_user_note) " &
                    "VALUES (" &
                    ":wf_file_name, :well_name, :well_contractor, :wf_barcode, :wf_title, " &
                    ":wf_authors, :wf_date, :wf_type, :wf_subject, :wf_group, " &
                    ":wf_num_of_page, :wf_note, :wf_doc_ver, :wf_file_path, :wf_file_size, " &
                    ":wf_load_by, SYSDATE, :wf_verified_by, :wf_verified_date, :approval_group, " &
                    ":approval_status, :approval_inameta_by, :approval_inameta_date, :approval_inameta_note, " &
                    ":approval_user_by, :approval_user_date, :approval_user_note)"

                    Using cmd As New OracleCommand(query, conn)
                        With cmd.Parameters
                            .Add(":wf_file_name", OracleDbType.Varchar2).Value = If(row.Cells("wf_file_name").Value, DBNull.Value)
                            .Add(":well_name", OracleDbType.Varchar2).Value = If(row.Cells("well_name").Value, DBNull.Value)
                            .Add(":well_contractor", OracleDbType.Varchar2).Value = If(row.Cells("well_contractor").Value, DBNull.Value)
                            .Add(":wf_barcode", OracleDbType.Varchar2).Value = If(row.Cells("wf_barcode").Value, DBNull.Value)
                            .Add(":wf_title", OracleDbType.Varchar2).Value = If(row.Cells("wf_title").Value, DBNull.Value)
                            .Add(":wf_authors", OracleDbType.Varchar2).Value = If(row.Cells("wf_authors").Value, DBNull.Value)
                            .Add(":wf_date", OracleDbType.Date).Value = If(IsDate(row.Cells("wf_date").Value), CDate(row.Cells("wf_date").Value), DBNull.Value)
                            .Add(":wf_type", OracleDbType.Varchar2).Value = If(row.Cells("wf_type").Value, DBNull.Value)
                            .Add(":wf_subject", OracleDbType.Varchar2).Value = If(row.Cells("wf_subject").Value, DBNull.Value)
                            .Add(":wf_group", OracleDbType.Varchar2).Value = If(row.Cells("wf_group").Value, DBNull.Value)
                            .Add(":wf_num_of_page", OracleDbType.Int32).Value = If(IsNumeric(row.Cells("wf_num_of_page").Value), CInt(row.Cells("wf_num_of_page").Value), DBNull.Value)
                            .Add(":wf_note", OracleDbType.Varchar2).Value = If(row.Cells("wf_note").Value, DBNull.Value)
                            .Add(":wf_doc_ver", OracleDbType.Varchar2).Value = If(row.Cells("wf_doc_ver").Value, DBNull.Value)
                            .Add(":wf_file_path", OracleDbType.Varchar2).Value = If(row.Cells("wf_file_path").Value, DBNull.Value)
                            .Add(":wf_file_size", OracleDbType.Int32).Value = If(IsNumeric(row.Cells("wf_file_size").Value), CInt(row.Cells("wf_file_size").Value), DBNull.Value)
                            .Add(":wf_load_by", OracleDbType.Varchar2).Value = If(row.Cells("wf_load_by").Value, DBNull.Value)
                            .Add(":wf_verified_by", OracleDbType.Varchar2).Value = If(row.Cells("wf_verified_by").Value, DBNull.Value)
                            .Add(":wf_verified_date", OracleDbType.Date).Value = If(IsDate(row.Cells("wf_verified_date").Value), CDate(row.Cells("wf_verified_date").Value), DBNull.Value)
                            .Add(":approval_group", OracleDbType.Varchar2).Value = If(row.Cells("approval_group").Value, DBNull.Value)
                            .Add(":approval_status", OracleDbType.Int32).Value = If(IsNumeric(row.Cells("approval_status").Value), CInt(row.Cells("approval_status").Value), DBNull.Value)
                            .Add(":approval_inameta_by", OracleDbType.Varchar2).Value = If(row.Cells("approval_inameta_by").Value, DBNull.Value)
                            .Add(":approval_inameta_date", OracleDbType.Date).Value = If(IsDate(row.Cells("approval_inameta_date").Value), CDate(row.Cells("approval_inameta_date").Value), DBNull.Value)
                            .Add(":approval_inameta_note", OracleDbType.Varchar2).Value = If(row.Cells("approval_inameta_note").Value, DBNull.Value)
                            .Add(":approval_user_by", OracleDbType.Varchar2).Value = If(row.Cells("approval_user_by").Value, DBNull.Value)
                            .Add(":approval_user_date", OracleDbType.Date).Value = If(IsDate(row.Cells("approval_user_date").Value), CDate(row.Cells("approval_user_date").Value), DBNull.Value)
                            .Add(":approval_user_note", OracleDbType.Varchar2).Value = If(row.Cells("approval_user_note").Value, DBNull.Value)
                        End With

                        cmd.ExecuteNonQuery()
                    End Using
                Case "WELL_LOG_DATA"
                    ' Ambil nilai terbaru untuk well_log_s
                    Dim wellLogS As Integer = getWellLogS() + 1

                    ' Step 1: Insert ke WELL_LOG
                    Dim queryLog As String = "INSERT INTO WELL_LOG (" &
                    "WELL_LOG_S, WL_PRODUCERS, WELL_NAME, WELL_CONTRACTOR, WL_LOG_TYPE, " &
                    "WL_RUN_NO, WL_RUN_DATE, WL_TOP_DEPTH, WL_BOTTOM_DEPTH, WL_DEPTH_U, " &
                    "WL_REMARKS, WL_CURVE_TYPE, WL_NOTE, WL_TITLE) VALUES (" &
                    ":WELL_LOG_S, :WL_PRODUCERS, :WELL_NAME, :WELL_CONTRACTOR, :WL_LOG_TYPE, " &
                    ":WL_RUN_NO, :WL_RUN_DATE, :WL_TOP_DEPTH, :WL_BOTTOM_DEPTH, :WL_DEPTH_U, " &
                    ":WL_REMARKS, :WL_CURVE_TYPE, :WL_NOTE, :WL_TITLE)"

                    Using cmdLog As New OracleCommand(queryLog, conn)
                        With cmdLog.Parameters
                            .Add(":WELL_LOG_S", OracleDbType.Int32).Value = wellLogS
                            .Add(":WL_PRODUCERS", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_producers").Value), DBNull.Value, row.Cells("wl_producers").Value)
                            .Add(":WELL_NAME", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("well_name").Value), DBNull.Value, row.Cells("well_name").Value)
                            .Add(":WELL_CONTRACTOR", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("well_contractor").Value), DBNull.Value, row.Cells("well_contractor").Value)
                            .Add(":WL_LOG_TYPE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_log_type").Value), DBNull.Value, row.Cells("wl_log_type").Value)
                            .Add(":WL_RUN_NO", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_run_no").Value), DBNull.Value, row.Cells("wl_run_no").Value)
                            .Add(":WL_RUN_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("wl_run_date").Value), CDate(row.Cells("wl_run_date").Value), DBNull.Value)
                            .Add(":WL_TOP_DEPTH", OracleDbType.Double).Value = If(IsNumeric(row.Cells("wl_top_depth").Value), CDbl(row.Cells("wl_top_depth").Value), DBNull.Value)
                            .Add(":WL_BOTTOM_DEPTH", OracleDbType.Double).Value = If(IsNumeric(row.Cells("wl_bottom_depth").Value), CDbl(row.Cells("wl_bottom_depth").Value), DBNull.Value)
                            .Add(":WL_DEPTH_U", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_depth_u").Value), DBNull.Value, row.Cells("wl_depth_u").Value)
                            .Add(":WL_REMARKS", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_remarks").Value), DBNull.Value, row.Cells("wl_remarks").Value)
                            .Add(":WL_CURVE_TYPE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_curve_type").Value), DBNull.Value, row.Cells("wl_curve_type").Value)
                            .Add(":WL_NOTE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_note").Value), DBNull.Value, row.Cells("wl_note").Value)
                            .Add(":WL_TITLE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_title").Value), DBNull.Value, row.Cells("wl_title").Value)
                        End With
                        cmdLog.ExecuteNonQuery()
                    End Using

                    ' Step 2: Insert ke WELL_LOG_DATA
                    Dim queryData As String = "INSERT INTO WELL_LOG_DATA (" &
                    "WLD_FILE_NAME, WLD_FILE_SIZE, WLD_FILE_PATH, WLD_NOTE, WLD_VERIFIED_BY, " &
                    "WLD_VERIFIED_DATE, WLD_LOAD_BY, WLD_LOAD_DATE, WELL_LOG_S, APPROVAL_GROUP, " &
                    "APPROVAL_STATUS, APPROVAL_INAMETA_BY, APPROVAL_INAMETA_DATE, APPROVAL_INAMETA_NOTE, " &
                    "APPROVAL_USER_BY, APPROVAL_USER_DATE, APPROVAL_USER_NOTE) VALUES (" &
                    ":WLD_FILE_NAME, :WLD_FILE_SIZE, :WLD_FILE_PATH, :WLD_NOTE, :WLD_VERIFIED_BY, " &
                    ":WLD_VERIFIED_DATE, :WLD_LOAD_BY, :WLD_LOAD_DATE, :WELL_LOG_S, :APPROVAL_GROUP, " &
                    ":APPROVAL_STATUS, :APPROVAL_INAMETA_BY, :APPROVAL_INAMETA_DATE, :APPROVAL_INAMETA_NOTE, " &
                    ":APPROVAL_USER_BY, :APPROVAL_USER_DATE, :APPROVAL_USER_NOTE)"

                    Using cmdData As New OracleCommand(queryData, conn)
                        With cmdData.Parameters
                            .Add(":WLD_FILE_NAME", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wld_file_name").Value), DBNull.Value, row.Cells("wld_file_name").Value)
                            .Add(":WLD_FILE_SIZE", OracleDbType.Int32).Value = If(IsNumeric(row.Cells("wld_file_size").Value), CInt(row.Cells("wld_file_size").Value), DBNull.Value)
                            .Add(":WLD_FILE_PATH", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wld_file_path").Value), DBNull.Value, row.Cells("wld_file_path").Value)
                            .Add(":WLD_NOTE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wld_note").Value), DBNull.Value, row.Cells("wld_note").Value)
                            .Add(":WLD_VERIFIED_BY", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wld_verified_by").Value), DBNull.Value, row.Cells("wld_verified_by").Value)
                            .Add(":WLD_VERIFIED_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("wld_verified_date").Value), CDate(row.Cells("wld_verified_date").Value), DBNull.Value)
                            .Add(":WLD_LOAD_BY", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wld_load_by").Value), DBNull.Value, row.Cells("wld_load_by").Value)
                            .Add(":WLD_LOAD_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("wld_load_date").Value), CDate(row.Cells("wld_load_date").Value), DBNull.Value)
                            .Add(":WELL_LOG_S", OracleDbType.Int32).Value = wellLogS
                            .Add(":APPROVAL_GROUP", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_group").Value), DBNull.Value, row.Cells("approval_group").Value)
                            .Add(":APPROVAL_STATUS", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_status").Value), DBNull.Value, row.Cells("approval_status").Value)
                            .Add(":APPROVAL_INAMETA_BY", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_inameta_by").Value), DBNull.Value, row.Cells("approval_inameta_by").Value)
                            .Add(":APPROVAL_INAMETA_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("approval_inameta_date").Value), CDate(row.Cells("approval_inameta_date").Value), DBNull.Value)
                            .Add(":APPROVAL_INAMETA_NOTE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_inameta_note").Value), DBNull.Value, row.Cells("approval_inameta_note").Value)
                            .Add(":APPROVAL_USER_BY", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_user_by").Value), DBNull.Value, row.Cells("approval_user_by").Value)
                            .Add(":APPROVAL_USER_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("approval_user_date").Value), CDate(row.Cells("approval_user_date").Value), DBNull.Value)
                            .Add(":APPROVAL_USER_NOTE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_user_note").Value), DBNull.Value, row.Cells("approval_user_note").Value)
                            '.Add(":WLD_FILE_NAME", OracleDbType.Varchar2).Value = If(row.Cells("wld_file_name").Value, DBNull.Value)
                            '.Add(":WLD_FILE_SIZE", OracleDbType.Int32).Value = If(IsNumeric(row.Cells("wld_file_size").Value), CInt(row.Cells("wld_file_size").Value), DBNull.Value)
                            '.Add(":WLD_FILE_PATH", OracleDbType.Varchar2).Value = If(row.Cells("wld_file_path").Value, DBNull.Value)
                            '.Add(":WLD_NOTE", OracleDbType.Varchar2).Value = If(row.Cells("wld_note").Value, DBNull.Value)
                            '.Add(":WLD_VERIFIED_BY", OracleDbType.Varchar2).Value = If(row.Cells("wld_verified_by").Value, DBNull.Value)
                            '.Add(":WLD_VERIFIED_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("wld_verified_date").Value), CDate(row.Cells("wld_verified_date").Value), DBNull.Value)
                            '.Add(":WLD_LOAD_BY", OracleDbType.Varchar2).Value = If(row.Cells("wld_load_by").Value, DBNull.Value)
                            '.Add(":WLD_LOAD_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("wld_load_date").Value), CDate(row.Cells("wld_load_date").Value), DBNull.Value)
                            '.Add(":WELL_LOG_S", OracleDbType.Int32).Value = wellLogS
                            '.Add(":APPROVAL_GROUP", OracleDbType.Varchar2).Value = If(row.Cells("approval_group").Value, DBNull.Value)
                            '.Add(":APPROVAL_STATUS", OracleDbType.Varchar2).Value = If(row.Cells("approval_status").Value, DBNull.Value)
                            '.Add(":APPROVAL_INAMETA_BY", OracleDbType.Varchar2).Value = If(row.Cells("approval_inameta_by").Value, DBNull.Value)
                            '.Add(":APPROVAL_INAMETA_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("approval_inameta_date").Value), CDate(row.Cells("approval_inameta_date").Value), DBNull.Value)
                            '.Add(":APPROVAL_INAMETA_NOTE", OracleDbType.Varchar2).Value = If(row.Cells("approval_inameta_note").Value, DBNull.Value)
                            '.Add(":APPROVAL_USER_BY", OracleDbType.Varchar2).Value = If(row.Cells("approval_user_by").Value, DBNull.Value)
                            '.Add(":APPROVAL_USER_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("approval_user_date").Value), CDate(row.Cells("approval_user_date").Value), DBNull.Value)
                            '.Add(":APPROVAL_USER_NOTE", OracleDbType.Varchar2).Value = If(row.Cells("approval_user_note").Value, DBNull.Value)
                        End With
                        cmdData.ExecuteNonQuery()
                    End Using
                Case "WELL_LOG_IMAGE"
                    ' Ambil nilai terbaru untuk well_log_s
                    Dim wellLogS As Integer = getWellLogS() + 1

                    ' Step 1: Insert ke WELL_LOG
                    Dim queryLog As String = "INSERT INTO WELL_LOG (" &
                    "WELL_LOG_S, WL_PRODUCERS, WELL_NAME, WELL_CONTRACTOR, WL_LOG_TYPE, " &
                    "WL_RUN_NO, WL_RUN_DATE, WL_TOP_DEPTH, WL_BOTTOM_DEPTH, WL_DEPTH_U, " &
                    "WL_REMARKS, WL_CURVE_TYPE, WL_NOTE, WL_TITLE) VALUES (" &
                    ":WELL_LOG_S, :WL_PRODUCERS, :WELL_NAME, :WELL_CONTRACTOR, :WL_LOG_TYPE, " &
                    ":WL_RUN_NO, :WL_RUN_DATE, :WL_TOP_DEPTH, :WL_BOTTOM_DEPTH, :WL_DEPTH_U, " &
                    ":WL_REMARKS, :WL_CURVE_TYPE, :WL_NOTE, :WL_TITLE)"

                    Using cmdLog As New OracleCommand(queryLog, conn)
                        With cmdLog.Parameters
                            .Add(":WELL_LOG_S", OracleDbType.Int32).Value = wellLogS
                            .Add(":WL_PRODUCERS", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_producers").Value), DBNull.Value, row.Cells("wl_producers").Value)
                            .Add(":WELL_NAME", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("well_name").Value), DBNull.Value, row.Cells("well_name").Value)
                            .Add(":WELL_CONTRACTOR", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("well_contractor").Value), DBNull.Value, row.Cells("well_contractor").Value)
                            .Add(":WL_LOG_TYPE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_log_type").Value), DBNull.Value, row.Cells("wl_log_type").Value)
                            .Add(":WL_RUN_NO", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_run_no").Value), DBNull.Value, row.Cells("wl_run_no").Value)
                            .Add(":WL_RUN_DATE", OracleDbType.Date).Value = If(IsDate(row.Cells("wl_run_date").Value), CDate(row.Cells("wl_run_date").Value), DBNull.Value)
                            .Add(":WL_TOP_DEPTH", OracleDbType.Double).Value = If(IsNumeric(row.Cells("wl_top_depth").Value), CDbl(row.Cells("wl_top_depth").Value), DBNull.Value)
                            .Add(":WL_BOTTOM_DEPTH", OracleDbType.Double).Value = If(IsNumeric(row.Cells("wl_bottom_depth").Value), CDbl(row.Cells("wl_bottom_depth").Value), DBNull.Value)
                            .Add(":WL_DEPTH_U", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_depth_u").Value), DBNull.Value, row.Cells("wl_depth_u").Value)
                            .Add(":WL_REMARKS", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_remarks").Value), DBNull.Value, row.Cells("wl_remarks").Value)
                            .Add(":WL_CURVE_TYPE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_curve_type").Value), DBNull.Value, row.Cells("wl_curve_type").Value)
                            .Add(":WL_NOTE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_note").Value), DBNull.Value, row.Cells("wl_note").Value)
                            .Add(":WL_TITLE", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wl_title").Value), DBNull.Value, row.Cells("wl_title").Value)
                        End With
                        cmdLog.ExecuteNonQuery()
                    End Using

                    ' Step 2: Insert ke WELL_LOG_IMAGE
                    Dim queryImage As String = "INSERT INTO WELL_LOG_IMAGE (" &
                    "WLI_FILE_NAME, WLI_FILE_SIZE, WLI_FILE_PATH, WLI_NOTE, WLI_VERIFIED_BY, " &
                    "WLI_VERIFIED_DATE, WLI_LOAD_BY, WLI_LOAD_DATE, WLI_VERTICAL_SCALE, WELL_LOG_S, " &
                    "WLI_HDR_FILE_NAME, WLI_HDR_FILE_PATH, WLI_HDR_FILE_SIZE, WLI_BARCODE, " &
                    "APPROVAL_GROUP, APPROVAL_STATUS, APPROVAL_INAMETA_BY, APPROVAL_INAMETA_DATE, " &
                    "APPROVAL_INAMETA_NOTE, APPROVAL_USER_BY, APPROVAL_USER_DATE, APPROVAL_USER_NOTE) VALUES (" &
                    ":WLI_FILE_NAME, :WLI_FILE_SIZE, :WLI_FILE_PATH, :WLI_NOTE, :WLI_VERIFIED_BY, " &
                    ":WLI_VERIFIED_DATE, :WLI_LOAD_BY, :WLI_LOAD_DATE, :WLI_VERTICAL_SCALE, :WELL_LOG_S, " &
                    ":WLI_HDR_FILE_NAME, :WLI_HDR_FILE_PATH, :WLI_HDR_FILE_SIZE, :WLI_BARCODE, " &
                    ":APPROVAL_GROUP, :APPROVAL_STATUS, :APPROVAL_INAMETA_BY, :APPROVAL_INAMETA_DATE, " &
                    ":APPROVAL_INAMETA_NOTE, :APPROVAL_USER_BY, :APPROVAL_USER_DATE, :APPROVAL_USER_NOTE)"

                    Using cmdImage As New OracleCommand(queryImage, conn)
                        With cmdImage.Parameters
                            .Add(":wli_file_name", OracleDbType.Varchar2).Value = row.Cells("wli_file_name").Value
                            .Add(":wli_file_size", OracleDbType.Int32).Value = If(IsDBNull(row.Cells("wli_file_size").Value), DBNull.Value, row.Cells("wli_file_size").Value)
                            .Add(":wli_file_path", OracleDbType.Varchar2).Value = row.Cells("wli_file_path").Value
                            .Add(":wli_note", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_note").Value), DBNull.Value, row.Cells("wli_note").Value)
                            .Add(":wli_verified_by", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_verified_by").Value), DBNull.Value, row.Cells("wli_verified_by").Value)
                            .Add(":wli_verified_date", OracleDbType.Date).Value = If(IsDBNull(row.Cells("wli_verified_date").Value), DBNull.Value, row.Cells("wli_verified_date").Value)
                            .Add(":wli_load_by", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_load_by").Value), DBNull.Value, row.Cells("wli_load_by").Value)
                            .Add(":wli_load_date", OracleDbType.Date).Value = If(IsDBNull(row.Cells("wli_load_date").Value), DBNull.Value, row.Cells("wli_load_date").Value)
                            .Add(":wli_vertical_scale", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_vertical_scale").Value), DBNull.Value, row.Cells("wli_vertical_scale").Value)
                            .Add(":well_log_s", OracleDbType.Int32).Value = wellLogS
                            .Add(":wli_hdr_file_name", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_hdr_file_name").Value), DBNull.Value, row.Cells("wli_hdr_file_name").Value)
                            .Add(":wli_hdr_file_path", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_hdr_file_path").Value), DBNull.Value, row.Cells("wli_hdr_file_path").Value)
                            .Add(":wli_hdr_file_size", OracleDbType.Int32).Value = If(IsDBNull(row.Cells("wli_hdr_file_size").Value), DBNull.Value, row.Cells("wli_hdr_file_size").Value)
                            .Add(":wli_barcode", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("wli_barcode").Value), DBNull.Value, row.Cells("wli_barcode").Value)
                            .Add(":approval_group", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_group").Value), DBNull.Value, row.Cells("approval_group").Value)
                            .Add(":approval_status", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_status").Value), DBNull.Value, row.Cells("approval_status").Value)
                            .Add(":approval_inameta_by", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_inameta_by").Value), DBNull.Value, row.Cells("approval_inameta_by").Value)
                            .Add(":approval_inameta_date", OracleDbType.Date).Value = If(IsDBNull(row.Cells("approval_inameta_date").Value), DBNull.Value, row.Cells("approval_inameta_date").Value)
                            .Add(":approval_inameta_note", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_inameta_note").Value), DBNull.Value, row.Cells("approval_inameta_note").Value)
                            .Add(":approval_user_by", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_user_by").Value), DBNull.Value, row.Cells("approval_user_by").Value)
                            .Add(":approval_user_date", OracleDbType.Date).Value = If(IsDBNull(row.Cells("approval_user_date").Value), DBNull.Value, row.Cells("approval_user_date").Value)
                            .Add(":approval_user_note", OracleDbType.Varchar2).Value = If(IsDBNull(row.Cells("approval_user_note").Value), DBNull.Value, row.Cells("approval_user_note").Value)
                            '.Add(":wli_file_name", OracleDbType.Varchar2).Value = row.Cells("wli_file_name").Value
                            '.Add(":wli_file_size", OracleDbType.Int32).Value = row.Cells("wli_file_size").Value
                            '.Add(":wli_file_path", OracleDbType.Varchar2).Value = row.Cells("wli_file_path").Value
                            '.Add(":wli_note", OracleDbType.Varchar2).Value = row.Cells("wli_note").Value
                            '.Add(":wli_verified_by", OracleDbType.Varchar2).Value = row.Cells("wli_verified_by").Value
                            '.Add(":wli_verified_date", OracleDbType.Date).Value = row.Cells("wli_verified_date").Value
                            '.Add(":wli_load_by", OracleDbType.Varchar2).Value = row.Cells("wli_load_by").Value
                            '.Add(":wli_load_date", OracleDbType.Date).Value = row.Cells("wli_load_date").Value
                            '.Add(":wli_vertical_scale", OracleDbType.Varchar2).Value = row.Cells("wli_vertical_scale").Value
                            '.Add(":well_log_s", OracleDbType.Int32).Value = wellLogS
                            '.Add(":wli_hdr_file_name", OracleDbType.Varchar2).Value = row.Cells("wli_hdr_file_name").Value
                            '.Add(":wli_hdr_file_path", OracleDbType.Varchar2).Value = row.Cells("wli_hdr_file_path").Value
                            '.Add(":wli_hdr_file_size", OracleDbType.Int32).Value = row.Cells("wli_hdr_file_size").Value
                            '.Add(":wli_barcode", OracleDbType.Varchar2).Value = row.Cells("wli_barcode").Value
                            '.Add(":approval_group", OracleDbType.Varchar2).Value = row.Cells("approval_group").Value
                            '.Add(":approval_status", OracleDbType.Varchar2).Value = row.Cells("approval_status").Value
                            '.Add(":approval_inameta_by", OracleDbType.Varchar2).Value = row.Cells("approval_inameta_by").Value
                            '.Add(":approval_inameta_date", OracleDbType.Date).Value = row.Cells("approval_inameta_date").Value
                            '.Add(":approval_inameta_note", OracleDbType.Varchar2).Value = row.Cells("approval_inameta_note").Value
                            '.Add(":approval_user_by", OracleDbType.Varchar2).Value = row.Cells("approval_user_by").Value
                            '.Add(":approval_user_date", OracleDbType.Date).Value = row.Cells("approval_user_date").Value
                            '.Add(":approval_user_note", OracleDbType.Varchar2).Value = row.Cells("approval_user_note").Value
                        End With
                        cmdImage.ExecuteNonQuery()
                    End Using
                Case "WELL_MASTER_LOG"
                    Dim query As String = "INSERT INTO well_master_log (" &
                    "wml_file_name, wml_file_size, wml_file_path, wml_title, wml_note, " &
                    "wml_date, wml_verified_by, wml_verified_date, wml_load_by, wml_doc_ver, " &
                    "well_name, well_contractor, wml_load_date, wml_barcode, wml_vertical_scale, " &
                    "wml_top_depth, wml_bottom_depth, wml_depth_u, approval_group) VALUES (" &
                    ":wml_file_name, :wml_file_size, :wml_file_path, :wml_title, :wml_note, " &
                    ":wml_date, :wml_verified_by, :wml_verified_date, :wml_load_by, :wml_doc_ver, " &
                    ":well_name, :well_contractor, SYSDATE, :wml_barcode, :wml_vertical_scale, " &
                    ":wml_top_depth, :wml_bottom_depth, :wml_depth_u, :approval_group)"

                    Using cmd As New OracleCommand(query, conn)
                        With cmd.Parameters
                            .Add(":wml_file_name", OracleDbType.Varchar2).Value = If(row.Cells("wml_file_name").Value, DBNull.Value)
                            .Add(":wml_file_size", OracleDbType.Int32).Value = If(IsNumeric(row.Cells("wml_file_size").Value), CInt(row.Cells("wml_file_size").Value), DBNull.Value)
                            .Add(":wml_file_path", OracleDbType.Varchar2).Value = If(row.Cells("wml_file_path").Value, DBNull.Value)
                            .Add(":wml_title", OracleDbType.Varchar2).Value = If(row.Cells("wml_title").Value, DBNull.Value)
                            .Add(":wml_note", OracleDbType.Varchar2).Value = If(row.Cells("wml_note").Value, DBNull.Value)
                            .Add(":wml_date", OracleDbType.Date).Value = If(IsDate(row.Cells("wml_date").Value), CDate(row.Cells("wml_date").Value), DBNull.Value)
                            .Add(":wml_verified_by", OracleDbType.Varchar2).Value = If(row.Cells("wml_verified_by").Value, DBNull.Value)
                            .Add(":wml_verified_date", OracleDbType.Date).Value = If(IsDate(row.Cells("wml_verified_date").Value), CDate(row.Cells("wml_verified_date").Value), DBNull.Value)
                            .Add(":wml_load_by", OracleDbType.Varchar2).Value = If(row.Cells("wml_load_by").Value, DBNull.Value)
                            .Add(":wml_doc_ver", OracleDbType.Varchar2).Value = If(row.Cells("wml_doc_ver").Value, DBNull.Value)
                            .Add(":well_name", OracleDbType.Varchar2).Value = If(row.Cells("well_name").Value, DBNull.Value)
                            .Add(":well_contractor", OracleDbType.Varchar2).Value = If(row.Cells("well_contractor").Value, DBNull.Value)
                            .Add(":wml_barcode", OracleDbType.Varchar2).Value = If(row.Cells("wml_barcode").Value, DBNull.Value)
                            .Add(":wml_vertical_scale", OracleDbType.Varchar2).Value = If(row.Cells("wml_vertical_scale").Value, DBNull.Value)
                            .Add(":wml_top_depth", OracleDbType.Double).Value = If(IsNumeric(row.Cells("wml_top_depth").Value), CDbl(row.Cells("wml_top_depth").Value), DBNull.Value)
                            .Add(":wml_bottom_depth", OracleDbType.Double).Value = If(IsNumeric(row.Cells("wml_bottom_depth").Value), CDbl(row.Cells("wml_bottom_depth").Value), DBNull.Value)
                            .Add(":wml_depth_u", OracleDbType.Varchar2).Value = If(row.Cells("wml_depth_u").Value, DBNull.Value)
                            .Add(":approval_group", OracleDbType.Varchar2).Value = If(row.Cells("approval_group").Value, DBNull.Value)
                        End With

                        cmd.ExecuteNonQuery()
                    End Using


                Case Else
                    Throw New Exception("Form tidak dikenali.")
                    Console.WriteLine("ini ke form dak kenal kami file222")
            End Select

        Catch ex As Exception
            MessageBox.Show("Insert DB gagal: " & ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Function CheckIfFtpFileExists(ByVal fileUri As String) As Boolean

        Dim request As FtpWebRequest = WebRequest.Create(fileUri)
        request.Credentials = New NetworkCredential(My.Settings.FTPUser, My.Settings.FTPPass)
        request.Method = WebRequestMethods.Ftp.GetFileSize
        request.Proxy = Nothing
        request.KeepAlive = False
        request.UseBinary = True

        Try
            Dim response As FtpWebResponse = request.GetResponse()
            ' THE FILE EXISTS
        Catch ex As WebException
            Dim response As FtpWebResponse = ex.Response
            If FtpStatusCode.ActionNotTakenFileUnavailable = response.StatusCode Then
                ' THE FILE DOES NOT EXIST
                Return False
            End If
        End Try

        Return True
    End Function

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnCancel.Click
        If bgWorker.WorkerSupportsCancellation Then
            fString = 1
            bgWorker.CancelAsync()
        End If

        _Del_BtnUpload(True)
        _Del_BtnCancel(False)
    End Sub

    Private Sub End_Excel_App(datestart As Date, dateEnd As Date)
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            If Process.StartTime >= datestart And Process.StartTime <= dateEnd Then
                Process.Kill()
                Exit For
            End If
        Next
    End Sub

    Private Sub txBrowse_TextChanged(sender As Object, e As EventArgs) Handles txBrowse.TextChanged
        If txBrowse IsNot Nothing Then
            browseStatus = 1
            cbForm.Enabled = True
        End If
    End Sub

    Private Sub txBrowseFile_TextChanged(sender As Object, e As EventArgs) Handles txBrowseFile.TextChanged
        If txBrowseFile IsNot Nothing Then
            browseStatus = 2
            cbForm.Enabled = True
        End If
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Help.Show()
    End Sub

    Private Sub GenerateColour()
        For colIndex As Integer = 0 To DataGridView1.Columns.Count - 1
            For rowIndex As Integer = 1 To DataGridView1.Rows.Count - 1
                If DataGridView1.Columns(colIndex).HeaderCell.Value = "FLAG" Then
                    'Default State
                    If DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = 0 Then

                        'Duplicate
                    ElseIf DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = 1 Then
                        DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Orange

                        'Upload Failed
                    ElseIf DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = 2 Then
                        DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Cyan

                        'Insert Failed
                    ElseIf DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = 3 Then
                        DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.OrangeRed

                        'Upload & Insert Success 
                    ElseIf DataGridView1.Rows(rowIndex - 1).Cells(colIndex).Value = 4 Then
                        DataGridView1.Rows(rowIndex - 1).DefaultCellStyle.BackColor = Color.Lime
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim strHeader As String = DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).HeaderCell.Value.ToString
        If strHeader.Contains("WF_TYPE") Or strHeader.Contains("RW_FTYPE") Then
            Dim selectedComboBox As ComboBox = DirectCast(e.Control, ComboBox)
            RemoveHandler selectedComboBox.SelectionChangeCommitted, AddressOf selectedComboBox_SelectionChangeCommitted
            AddHandler selectedComboBox.SelectionChangeCommitted, AddressOf selectedComboBox_SelectionChangeCommitted
        End If
    End Sub

    Private Sub selectedComboBox_SelectionChangeCommitted(ByVal sender As Object, ByVal e As EventArgs)
        Dim selectedCombobox As ComboBox = DirectCast(sender, ComboBox)
        If selectedCombobox.SelectedItem IsNot Nothing Then
            Dim drv As DataRowView = selectedCombobox.SelectedItem

            If DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).HeaderCell.Value = "WF_TYPE" Then
                cbValue = drv("r_document_type_id").ToString
                If cbValue = "PRE" Then
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(13).Value = "DIGDAT\WELLTEST"
                ElseIf cbValue = "PPA" Then
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(13).Value = "DIGDAT\WELLLOG\PETRAN"
                End If

            ElseIf DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).HeaderCell.Value = "RW_FTYPE" Then
                cbValue = drv("r_document_type_id").ToString
                If cbValue = "CPH" Then
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(4).Value = "DIGDAT\REPO\WELL\CORE"
                ElseIf cbValue = "PRS" Then
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(4).Value = "DIGDAT\REPO\WELL\SURVEY"
                End If
            End If
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.ColumnIndex <> -1 Then
            If sender.columns(e.ColumnIndex).headercell.value = "WF_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WLD_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WLI_FILE_NAME" Or sender.columns(e.ColumnIndex).headercell.value = "WLI_HDR_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WML_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "WC_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "GCI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "GRI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RW_FNAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RA_FNAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "SMI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "SPDD_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "PFM_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RMI_FILE_NAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            ElseIf sender.columns(e.ColumnIndex).headercell.value = "RA_FNAME" Then
                btnRename.Enabled = True
                btnOpenFile.Enabled = True
                activeCell = DataGridView1.CurrentCell
                btnOpenFile_Click(sender, e)
            End If
        End If
    End Sub

    Public Class SafeDataGridView
        Inherits DataGridView

        Public Property EnableRowHeaderDoubleClick As Boolean

        Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
            Try
                MyBase.OnPaint(e)
            Catch generatedExceptionName As Exception
                Me.Invalidate()
            End Try
        End Sub
    End Class

    Private Sub LogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogToolStripMenuItem.Click
        If File.Exists(filelogErr) = True Then
            Dim process As Process = New Process()
            Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
            startInfo.FileName = filelogErr
            process.StartInfo = startInfo
            process.Start()
        Else
            MessageBox.Show("File does Not exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.SelectedCells.Count = 1 Then
            DataGridView1.SelectAll()
        End If
    End Sub

    Private Sub DataGridView1_RowHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        If EnableRowHeaderDoubleClick = False Then Exit Sub
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Dim gridView As DataGridView = TryCast(sender, DataGridView)
        If gridView IsNot Nothing Then
            For Each r As DataGridViewRow In gridView.Rows
                gridView.Rows(r.Index).HeaderCell.Value = (r.Index + 1).ToString()
            Next
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If Me.RadioButton2.Checked = False And My.Settings.Mode = 2 Then
            Me.RadioButton1.Checked = True
            My.Settings.Mode = 1
            My.Settings.Save()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If Me.RadioButton1.Checked = False And My.Settings.Mode = 1 Then
            Me.RadioButton2.Checked = True
            My.Settings.Mode = 2
            My.Settings.Save()
            UNCLogin.ShowDialog()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim conn As New OracleConnection(conString)
        Try
            conn.Open()
            ' Lakukan perintah SQL di sini
            'MessageBox.Show("Koneksi berhasil!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            For Each row As DataGridViewRow In DataGridView1.Rows
                If TypeOf row.Cells("wf_date").Value Is Date Then
                    xx = "sudah"
                Else
                    xx = "belum"
                End If
                If row.IsNewRow Then Continue For
                relativeTargetPath = row.Cells("wf_file_path").Value.ToString().Trim()
                x = row.Cells("path_sourcefile").Value.ToString().Trim()
                wfauth = row.Cells("wf_authors").Value.ToString().Trim()
                px = My.Settings.DigdatHost & relativeTargetPath
                c = row.Cells("wf_file_name").Value.ToString().Trim()
                v = Path.Combine(x, c)
                'b = Path.Combine(targetFolder, c)

            Next
            MessageBox.Show("Debug Path:" & xx & Environment.NewLine &
                "pathSource: " & x & Environment.NewLine &
                "fileName: " & c & Environment.NewLine &
                "Full Path: " & v &
                "px: " & px &
                "wf = " & wfauth,
                "DEBUG", MessageBoxButtons.OK, MessageBoxIcon.Information)
            conn.Close()
        Catch ex As Exception
            MessageBox.Show("Koneksi gagal: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'For Each row As DataGridViewRow In DataGridView1.Rows
        '    If TypeOf row.Cells("wf_date").Value Is Date Then
        '        xx = "sudah"
        '    Else
        '        xx = "belum"
        '    End If
        '    If row.IsNewRow Then Continue For

        '    x = row.Cells("path_sourcefile").Value.ToString().Trim()
        '    px = My.Settings.DigdatHost & relativeTargetPath
        '    c = row.Cells("wf_file_name").Value.ToString().Trim()
        '    v = Path.Combine(x, c)
        '    'b = Path.Combine(targetFolder, c)

        'Next
        MessageBox.Show("Debug Path:" & xx & Environment.NewLine &
                "pathSource: " & x & Environment.NewLine &
                "fileName: " & c & Environment.NewLine &
                "Full Path: " & v &
                "px: " & px,
                "DEBUG", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'CopyFiles()
    End Sub

    Private Sub bgexe_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgexe.DoWork
        _Del_BtnBrowse(False)
        _Del_BtnBrowseFile(False)
        _Del_BtnExecute(False)
        _Del_BtnExport(False)
        _Del_BtnUpload(False)
        'MessageBox.Show("FTP User: " & My.Settings.FTPUser & vbCrLf &
        '        "Digdat Host: " & My.Settings.DigdatHost, "Konfigurasi FTP")


        CountErr = 0
        cbFormSelValue = cbForm.SelectedValue.ToString
        Dim strCompare As String = String.Empty

        If browseStatus = 1 Then
            If Trim(txBrowse.Text) <> "" Then
                If Directory.Exists(txBrowse.Text) Then

                    Dim types() As String = Split(txType.Text, ";")

                    If cbFindMod.Checked = True Then
                        files = Directory.GetFiles(txBrowse.Text, txFind.Text, SearchOption.AllDirectories)
                    Else
                        files = Directory.GetFiles(txBrowse.Text, txFind.Text, SearchOption.TopDirectoryOnly)
                    End If

                    Array.Sort(files)

                    Dim x As Integer

                    respon = MessageBox.Show(files.Length.ToString + " file(s) found, Do you want to continue ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    Select Case respon
                        Case vbYes

                            Dim xlApp As Excel.Application = New Excel.Application()
                            Dim xlWorkBook As Excel.Workbook
                            Dim xlWorkSheet As Excel.Worksheet
                            Dim misValue As Object = Reflection.Missing.Value
                            Dim dateStart As Date = Date.Now
                            Dim oldci As CultureInfo = Thread.CurrentThread.CurrentCulture
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-us")

                            If xlApp Is Nothing Then
                                MessageBox.Show("Excel is not properly installed!!")
                                Return
                            End If

                            If cbForm.SelectedIndex = 0 Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "Path"
                                xlWorkSheet.Cells(1, 2) = "File Name"
                            ElseIf cbForm.SelectedValue.ToString = "WELL_FILE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WF_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 3) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 4) = "WF_BARCODE"
                                xlWorkSheet.Cells(1, 5) = "WF_TITLE"
                                xlWorkSheet.Cells(1, 6) = "WF_AUTHORS"
                                xlWorkSheet.Cells(1, 7) = "WF_DATE"
                                xlWorkSheet.Cells(1, 8) = "WF_TYPE"
                                xlWorkSheet.Cells(1, 9) = "WF_SUBJECT"
                                xlWorkSheet.Cells(1, 10) = "WF_GROUP"
                                xlWorkSheet.Cells(1, 11) = "WF_NUM_OF_PAGE"
                                xlWorkSheet.Cells(1, 12) = "WF_NOTE"
                                xlWorkSheet.Cells(1, 13) = "WF_DOC_VER"
                                xlWorkSheet.Cells(1, 14) = "WF_FILE_PATH"
                                xlWorkSheet.Cells(1, 15) = "WF_FILE_SIZE"
                                xlWorkSheet.Cells(1, 16) = "WF_LOAD_BY"
                                xlWorkSheet.Cells(1, 17) = "WF_LOADED_DATE"
                                xlWorkSheet.Cells(1, 18) = "WF_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 19) = "WF_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 24) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 25) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 26) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 27) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 28) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 29) = "FLAG"
                                xlWorkSheet.Cells(1, 30) = "FLAG_UPLOAD"
                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "S1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_DATA" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WLD_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WLD_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WLD_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WLD_NOTE"
                                xlWorkSheet.Cells(1, 5) = "WLD_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 6) = "WLD_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 7) = "WLD_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "WLD_LOAD_DATE"
                                xlWorkSheet.Cells(1, 9) = "WELL_LOG_S"
                                xlWorkSheet.Cells(1, 10) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 11) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 12) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 13) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 14) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 18) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 19) = "FLAG"
                                xlWorkSheet.Cells(1, 20) = "FLAG_UPLOAD"
                                xlWorkSheet.Cells(1, 21) = "WL_PRODUCERS"
                                xlWorkSheet.Cells(1, 22) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 23) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 24) = "WL_LOG_TYPE"
                                xlWorkSheet.Cells(1, 25) = "WL_RUN_NO"
                                xlWorkSheet.Cells(1, 26) = "WL_RUN_DATE"
                                xlWorkSheet.Cells(1, 27) = "WL_TOP_DEPTH"
                                xlWorkSheet.Cells(1, 28) = "WL_BOTTOM_DEPTH"
                                xlWorkSheet.Cells(1, 29) = "WL_DEPTH_U"
                                xlWorkSheet.Cells(1, 30) = "WL_REMARKS"
                                xlWorkSheet.Cells(1, 31) = "WL_CURVE_TYPE"
                                xlWorkSheet.Cells(1, 32) = "WL_NOTE"
                                xlWorkSheet.Cells(1, 33) = "WL_TITLE"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_IMAGE" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WLI_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WLI_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WLI_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WLI_NOTE"
                                xlWorkSheet.Cells(1, 5) = "WLI_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 6) = "WLI_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 7) = "WLI_LOAD_BY"
                                xlWorkSheet.Cells(1, 8) = "WLI_LOAD_DATE"
                                xlWorkSheet.Cells(1, 9) = "WLI_VERTICAL_SCALE"
                                xlWorkSheet.Cells(1, 10) = "WELL_LOG_S"
                                xlWorkSheet.Cells(1, 11) = "WLI_HDR_FILE_NAME"
                                xlWorkSheet.Cells(1, 12) = "WLI_HDR_FILE_PATH"
                                xlWorkSheet.Cells(1, 13) = "WLI_HDR_FILE_SIZE"
                                xlWorkSheet.Cells(1, 14) = "WLI_BARCODE"
                                xlWorkSheet.Cells(1, 15) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 16) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 17) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 18) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 23) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 24) = "FLAG"
                                xlWorkSheet.Cells(1, 25) = "FLAG_UPLOAD"
                                xlWorkSheet.Cells(1, 26) = "WL_PRODUCERS"
                                xlWorkSheet.Cells(1, 27) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 28) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 29) = "WL_LOG_TYPE"
                                xlWorkSheet.Cells(1, 30) = "WL_RUN_NO"
                                xlWorkSheet.Cells(1, 31) = "WL_RUN_DATE"
                                xlWorkSheet.Cells(1, 32) = "WL_TOP_DEPTH"
                                xlWorkSheet.Cells(1, 33) = "WL_BOTTOM_DEPTH"
                                xlWorkSheet.Cells(1, 34) = "WL_DEPTH_U"
                                xlWorkSheet.Cells(1, 35) = "WL_REMARKS"
                                xlWorkSheet.Cells(1, 36) = "WL_CURVE_TYPE"
                                xlWorkSheet.Cells(1, 37) = "WL_NOTE"
                                xlWorkSheet.Cells(1, 38) = "WL_TITLE"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "T1")
                                    .Font.Bold = True
                                End With
                            ElseIf cbForm.SelectedValue.ToString = "WELL_MASTER_LOG" Then
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                xlWorkSheet.Cells(1, 1) = "WML_FILE_NAME"
                                xlWorkSheet.Cells(1, 2) = "WML_FILE_SIZE"
                                xlWorkSheet.Cells(1, 3) = "WML_FILE_PATH"
                                xlWorkSheet.Cells(1, 4) = "WML_TITLE"
                                xlWorkSheet.Cells(1, 5) = "WML_NOTE"
                                xlWorkSheet.Cells(1, 6) = "WML_DATE"
                                xlWorkSheet.Cells(1, 7) = "WML_VERIFIED_BY"
                                xlWorkSheet.Cells(1, 8) = "WML_VERIFIED_DATE"
                                xlWorkSheet.Cells(1, 9) = "WML_LOAD_BY"
                                xlWorkSheet.Cells(1, 10) = "WML_DOC_VER"
                                xlWorkSheet.Cells(1, 11) = "WELL_NAME"
                                xlWorkSheet.Cells(1, 12) = "WELL_CONTRACTOR"
                                xlWorkSheet.Cells(1, 13) = "WML_LOAD_DATE"
                                xlWorkSheet.Cells(1, 14) = "WML_BARCODE"
                                xlWorkSheet.Cells(1, 15) = "WML_VERTICAL_SCALE"
                                xlWorkSheet.Cells(1, 16) = "WML_TOP_DEPTH"
                                xlWorkSheet.Cells(1, 17) = "WML_BOTTOM_DEPTH"
                                xlWorkSheet.Cells(1, 18) = "WML_DEPTH_U"
                                xlWorkSheet.Cells(1, 19) = "APPROVAL_GROUP"
                                xlWorkSheet.Cells(1, 20) = "APPROVAL_STATUS"
                                xlWorkSheet.Cells(1, 21) = "APPROVAL_INAMETA_BY"
                                xlWorkSheet.Cells(1, 22) = "APPROVAL_INAMETA_DATE"
                                xlWorkSheet.Cells(1, 23) = "APPROVAL_INAMETA_NOTE"
                                xlWorkSheet.Cells(1, 24) = "APPROVAL_USER_BY"
                                xlWorkSheet.Cells(1, 25) = "APPROVAL_USER_DATE"
                                xlWorkSheet.Cells(1, 26) = "APPROVAL_USER_NOTE"
                                xlWorkSheet.Cells(1, 27) = "PATH_SOURCEFILE"
                                xlWorkSheet.Cells(1, 28) = "FLAG"
                                xlWorkSheet.Cells(1, 29) = "FLAG_UPLOAD"

                                ' Format A1:D1 as bold, vertical alignment = center.
                                With xlWorkSheet.Range("A1", "O1")
                                    .Font.Bold = True
                                End With
                            Else
                                xlApp = New Excel.ApplicationClass()
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                            End If

                            Array.Reverse(files)
                            CreateFile()
                            log(0, "File processed", (files.Length).ToString)

                            Dim rowXls As Integer = 1
                            For x = 0 To files.Length - 1
                                For Each typee As String In types
                                    Try
                                        If Path.GetFileName(files(x)).ToUpper.EndsWith(Trim(typee)) And Trim(typee) <> "" Then
                                            If Not IsNothing(xlApp) Then
                                                Dim fileinfo = New FileInfo(Path.GetFullPath(files(x)))
                                                Dim fileSize As Double = Math.Round(fileinfo.Length / 1024)

                                                If cbForm.SelectedIndex = 0 Then
                                                    xlWorkSheet.Cells(rowXls + 1, 1) = Path.GetDirectoryName(files(x).ToString.ToUpper)
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_FILE" Then
                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = getWell(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "SBS"
                                                    xlWorkSheet.Cells(rowXls + 1, 4) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 4).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 5) = getTitle(Path.GetFileName(files(x).ToString.ToUpper))
                                                    xlWorkSheet.Cells(rowXls + 1, 6) = getAuthor(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 6).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 7).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 7).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 7).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 8) = getWfType(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 8).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 8).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 9) = getWfSbj(Path.GetDirectoryName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 9).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    If (Path.GetFileName(files(x).ToString.ToUpper.EndsWith(".PDF"))) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 11) = getNumberOfPdfPages(Path.GetFullPath(files(x).ToString)).ToString
                                                        If fString <> 0 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 11).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 11).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            fString = 0
                                                        End If
                                                    End If

                                                    'Check Filename to be used for Destination Path
                                                    dataGroup = xlWorkSheet.Cells(rowXls + 1, 1).Value.IndexOf("PETRAN")
                                                    If (dataGroup > -1) Then
                                                        dataGroup = "GNG"
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLLOG\PETRAN"
                                                    ElseIf xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString.ToUpper.Contains("PENGUKURAN") And xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString.ToUpper.Contains("TEKANAN") And xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString.ToUpper.Contains("DASAR") Then
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLTEST"
                                                    Else
                                                        dataGroup = "PETRO"
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLREPORT\WELLREPORTIMAGE"
                                                    End If

                                                    If (xlWorkSheet.Cells(rowXls + 1, 8).value = "PRE" And xlWorkSheet.Cells(rowXls + 1, 9).value = "B3") And (Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).value)) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLTEST"
                                                    ElseIf (xlWorkSheet.Cells(rowXls + 1, 8).value = "PPA" And xlWorkSheet.Cells(rowXls + 1, 9).value = "D4") And (Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).value)) Then
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = "DIGDAT\WELLLOG\PETRAN"
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 15) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 16) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = getApprovalGroupName(xlWorkSheet.Cells(rowXls + 1, 2).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, dataGroup)
                                                    xlWorkSheet.Cells(rowXls + 1, 21) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 28) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = 0

                                                    If Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).Value) Then
                                                        'Check PK in database
                                                        CheckPKWellFile(xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 2).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 8).Value.ToString, xlWorkSheet.Cells(rowXls + 1, 9).Value.ToString)
                                                        If fString = 1 Then
                                                            xlWorkSheet.Cells(rowXls + 1, 29) = 1
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            xlWorkSheet.Cells(rowXls + 1, 2).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            xlWorkSheet.Cells(rowXls + 1, 8).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                        Else
                                                            xlWorkSheet.Cells(rowXls + 1, 29) = 0
                                                        End If
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 30) = 0

                                                    If Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 4).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 5).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 6).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 8).Value) And Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 9).Value) And xlWorkSheet.Cells(rowXls + 1, 29).value = 0 Then
                                                        xlWorkSheet.Cells(rowXls + 1, 30) = 1
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_DATA" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\WELLLOG\WELLLOGDATA"
                                                    xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 10) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                    xlWorkSheet.Cells(rowXls + 1, 18) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 19) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 20) = 1
                                                    xlWorkSheet.Cells(rowXls + 1, 22) = wellname
                                                    xlWorkSheet.Cells(rowXls + 1, 23) = "SBS"

                                                    xlWorkSheet.Cells(rowXls + 1, 25) = getRunNo(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 25).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 25).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 25).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 26) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 26).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 26).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 26).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = getElevetionUnit(wellname)
                                                    xlWorkSheet.Cells(rowXls + 1, 31) = getCtLas(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        ' Format
                                                        xlWorkSheet.Cells(rowXls + 1, 31).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        fString = 0
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_LOG_IMAGE" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    If Path.GetFileName(files(x).ToString.ToUpper).Contains("_HDR.") Then
                                                        'If rowXls > 1 Then
                                                        rowXls -= 1
                                                        'End If

                                                        xlWorkSheet.Cells(rowXls + 1, 11) = Path.GetFileName(files(x).ToString.ToUpper)
                                                        xlWorkSheet.Cells(rowXls + 1, 12) = "DIGDAT\WELLLOG\WELLLOGIMAGE_HDR"
                                                        xlWorkSheet.Cells(rowXls + 1, 13) = fileSize.ToString
                                                        rowXls += 1
                                                        Continue For
                                                    Else
                                                        xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                        xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\WELLLOG\WELLLOGIMAGE"
                                                        xlWorkSheet.Cells(rowXls + 1, 7) = getLoadBy(Login.TextBox1.Text)
                                                        xlWorkSheet.Cells(rowXls + 1, 9) = getScale(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 9).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 9).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 14) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 14).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 15) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                        xlWorkSheet.Cells(rowXls + 1, 23) = Path.GetDirectoryName(files(x).ToString)
                                                        xlWorkSheet.Cells(rowXls + 1, 24) = 0
                                                        xlWorkSheet.Cells(rowXls + 1, 25) = 1
                                                        xlWorkSheet.Cells(rowXls + 1, 27) = wellname
                                                        xlWorkSheet.Cells(rowXls + 1, 28) = "SBS"
                                                        xlWorkSheet.Cells(rowXls + 1, 30) = getRunNo(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 30).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 30).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 30).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If
                                                        xlWorkSheet.Cells(rowXls + 1, 31) = getDate(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            If fString = 1 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 31).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            ElseIf fString = 2 Then
                                                                ' Format
                                                                xlWorkSheet.Cells(rowXls + 1, 31).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                                xlWorkSheet.Cells(rowXls + 1, 31).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            End If
                                                            fString = 0
                                                        End If

                                                        xlWorkSheet.Cells(rowXls + 1, 34) = getElevetionUnit(wellname)

                                                        xlWorkSheet.Cells(rowXls + 1, 36) = getCtTif(Path.GetFileName(files(x).ToString.ToUpper))
                                                        If fString <> 0 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 36).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                            fString = 0
                                                        End If

                                                        If Not String.IsNullOrEmpty(xlWorkSheet.Cells(rowXls + 1, 1).Value) Then
                                                            'Check PK in database
                                                            CheckPKWellLogImage(xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString)
                                                            If fString = 1 Then
                                                                xlWorkSheet.Cells(rowXls + 1, 24) = 1
                                                                xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Orange)
                                                            End If
                                                        End If
                                                    End If

                                                ElseIf cbForm.SelectedValue.ToString = "WELL_MASTER_LOG" Then
                                                    Dim wellname As String = getWell(Path.GetFileName(files(x).ToString.ToUpper))

                                                    xlWorkSheet.Cells(rowXls + 1, 1) = getCheck(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 1).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If
                                                    xlWorkSheet.Cells(rowXls + 1, 2) = fileSize.ToString
                                                    xlWorkSheet.Cells(rowXls + 1, 3) = "DIGDAT\WELLLOG\MASTERLOG"
                                                    xlWorkSheet.Cells(rowXls + 1, 9) = getLoadBy(Login.TextBox1.Text)
                                                    xlWorkSheet.Cells(rowXls + 1, 11) = wellname
                                                    xlWorkSheet.Cells(rowXls + 1, 12) = "SBS"

                                                    xlWorkSheet.Cells(rowXls + 1, 14) = getBarcode(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 14).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 14).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 15) = getScale(Path.GetFileName(files(x).ToString.ToUpper))
                                                    If fString <> 0 Then
                                                        If fString = 1 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        ElseIf fString = 2 Then
                                                            ' Format
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Font.Color = ColorTranslator.ToOle(Color.Red)
                                                            xlWorkSheet.Cells(rowXls + 1, 15).Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                                                        End If
                                                        fString = 0
                                                    End If

                                                    xlWorkSheet.Cells(rowXls + 1, 18) = getElevetionUnit(wellname)
                                                    xlWorkSheet.Cells(rowXls + 1, 19) = getApprovalGroupName(wellname, xlWorkSheet.Cells(rowXls + 1, 1).Value.ToString, "GNG")
                                                    xlWorkSheet.Cells(rowXls + 1, 27) = Path.GetDirectoryName(files(x).ToString)
                                                    xlWorkSheet.Cells(rowXls + 1, 28) = 0
                                                    xlWorkSheet.Cells(rowXls + 1, 29) = 1
                                                End If
                                                rowXls += 1
                                            End If
                                        End If
                                    Catch ex As Exception
                                        log(x + 1, "Error", Path.GetFileName(files(x).ToString) + ";" + ex.Message.ToString)
                                        CountErr += 1
                                    End Try
                                Next
                                ToolStripProgressBar((100 / files.Length) * (x + 1 - CountErr))
                                ToolStripStatusLabelTxt2("File " & x + 1 - CountErr & " of " & files.Length - CountErr)
                            Next

                            'ToolStripStatusLabelTxt2("Complete")

                            If Not IsNothing(xlApp) Then
                                Dim appDom As String = AppDomain.CurrentDomain.BaseDirectory + "output\"
                                xlApp.DisplayAlerts = False
                                Try
                                    xlWorkBook.SaveAs(appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + cbForm.SelectedValue.ToString + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)
                                End Try
                                xlWorkBook.Close(True, misValue, misValue)
                                xlApp.DisplayAlerts = True
                                xlApp.Quit()
                                Dim dateEnd As Date = Date.Now
                                End_Excel_App(dateStart, dateEnd)

                                releaseObject(xlWorkSheet)
                                releaseObject(xlWorkBook)
                                releaseObject(xlApp)

                                'Create listview
                                Dim filePath As String = appDom + "" + "GLF_" + Date.Now.Year.ToString + bln + tgll + "_" + cbForm.SelectedValue.ToString + ".xls"

                                Try
                                    FillDataGriedView(filePath)
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End Try
                            End If

                            If CountErr > 0 Then
                                MessageBox.Show("Searching complete, " + files.Length.ToString + " Files found, With " + CountErr.ToString + " Error Found" + Environment.NewLine + "Please ensure that you fill the fields which highlight by yellow color!", "Finish", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Else
                                MessageBox.Show("Searching complete, " + files.Length.ToString + " Files found" + Environment.NewLine + "Please ensure that you fill the fields which highlight by yellow color!", "Finish", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If

                    End Select
                Else
                    MessageBox.Show("Directory does not exist", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Else
                MessageBox.Show("Directory cannot be null", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        ElseIf browseStatus = 2 Then
            respon = MessageBox.Show("Are you sure ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            Select Case respon
                Case vbYes
                    Try
                        FillDataGriedView(txBrowseFile.Text)
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
            End Select
        End If

        formType = cbForm.SelectedValue
        ToolStripStatusLabelTxt1("Ready")

        btnBrowse.Enabled = True
        btnBrowseFile.Enabled = True
        btnExecute.Enabled = True
        btnExport.Enabled = True
        btnUpload.Enabled = True
    End Sub
End Class