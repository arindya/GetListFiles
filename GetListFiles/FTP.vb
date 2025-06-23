Imports System.Net
Imports System.IO

Public Class FTP
    Inherits GLV

    Public Sub UploadFiles(ByVal sourceFile As String, ByVal targetFile As String)
        targetFile = targetFile.Replace("DIGDAT", "").Replace("\", "/")
        targetFile = My.Settings.FTPHost + targetFile

        Try
            Dim request As FtpWebRequest = (FtpWebRequest).Create(targetFile + "/" + Path.GetFileName(sourceFile))
            request.Credentials = New NetworkCredential(My.Settings.FTPUser, My.Settings.FTPPass)
            request.Proxy = Nothing
            request.KeepAlive = False
            request.Method = WebRequestMethods.Ftp.UploadFile
            request.UseBinary = True

            Dim info As FileInfo = New FileInfo(sourceFile)
            request.ContentLength = info.Length

            'Create buffer for file contents
            Dim buffLength As Integer = 16384
            Dim buff(buffLength) As Byte
            Dim FileSize As Long = info.Length
            Dim FileSizeDescription As String = GetFileSize(FileSize).ToString
            Dim sentBytes As Long = 0

            'Upload file to FTP
            Try
                Using instream As FileStream = info.OpenRead()
                    Using outstream As Stream = request.GetRequestStream()
                        Dim bytesRead As Integer = instream.Read(buff, 0, buffLength)
                        While (bytesRead > 0)
                            outstream.Write(buff, 0, bytesRead)
                            bytesRead = instream.Read(buff, 0, buffLength)

                            sentBytes += bytesRead
                            Dim SummaryText As String = String.Format("Transferred {0} / {1}", GetFileSize(sentBytes), FileSizeDescription)
                            BackgroundWorker1.ReportProgress(Convert.ToInt32(Convert.ToDecimal(sentBytes) / Convert.ToDecimal(FileSize) * 100), SummaryText)
                        End While
                        outstream.Close()
                    End Using
                    instream.Close()
                End Using
            Catch e As WebException
                Dim status As FtpWebResponse = e.Response
                Throw
            End Try

            Dim response As FtpWebResponse = request.GetResponse()
            response.Close()
        Catch e As Exception
            Throw
        End Try
    End Sub

    Public Function GetFileSize(ByVal numBytes As Long) As String
        Dim fileSize As String = String.Empty

        If numBytes > 1073741824 Then
            fileSize = String.Format("{0:0.00} Gb", Convert.ToDouble(numBytes) / 1073741824)
        ElseIf (numBytes > 1048576)
            fileSize = String.Format("{0:0.00} Mb", Convert.ToDouble(numBytes) / 1048576)
        Else
            fileSize = String.Format("{0:0} Kb", Convert.ToDouble(numBytes) / 1024)
        End If

        If (fileSize = "0 Kb") Then
            fileSize = "1 Kb"
        End If
        Return fileSize
    End Function
End Class
