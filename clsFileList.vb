Imports System.IO
Imports System.IO.Directory
Imports System.IO.File
Public Class clsFileList
    Public Sub ListFiles(ByVal BaseDir As DirectoryInfo, ByVal sw As StreamWriter)
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                Dim Entry() As String = { _
                    di.FullName, _
                    di.Attributes.ToString, _
                    di.CreationTime.ToString, _
                    di.LastAccessTime.ToString, _
                    di.LastWriteTime.ToString}
                sw.WriteLine(String.Format("""{0}"",0 bytes,""{1}"",{2},{3},{4}", Entry))
                ListFiles(di, sw)
            Next
            Dim fiList As FileInfo() = BaseDir.GetFiles()
            For Each fi As FileInfo In fiList
                Dim Entry() As String = { _
                    fi.FullName, _
                    fi.Length.ToString("#,##0"), _
                    fi.Attributes.ToString, _
                    fi.CreationTime.ToString, _
                    fi.LastAccessTime.ToString, _
                    fi.LastWriteTime.ToString}
                sw.WriteLine(String.Format("""{0}"",""{1} bytes"",""{2}"",{3},{4},{5}", Entry))
            Next
        Catch ex As Exception
            sw.WriteLine(ex.ToString)
        End Try
    End Sub
    'Entry point which delegates to C-style main Private Function
    Public Overloads Shared Sub Main()
        System.Environment.ExitCode = Main(System.Environment.GetCommandLineArgs())
    End Sub
    Private Overloads Shared Function Main(ByVal args() As String) As Integer
        Dim sw As StreamWriter
        Dim OutputFileName As String
        Try
            Dim fl As New clsFileList
            Dim Entry() As String = { _
                "Path", _
                "Size (in bytes)", _
                "Attributes", _
                "CreationTime", _
                "LastAccessTime", _
                "LastWriteTime"}


            If args.Length < 2 Then
                Dim Message As String = _
                    "args(0) = Full path name of the executable" & vbCrLf & _
                    "args(1) = Root directory to list files" & vbCrLf & _
                    "args(2) = Optional output file name"
                MessageBox.Show(Message, "FileList", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If

            Dim Root As DirectoryInfo = New DirectoryInfo(args(1))
            If Not Root.Exists Then
                MessageBox.Show(String.Format("{0} does not exist!", args(1)), "FileList", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If

            If args.Length = 3 Then
                OutputFileName = args(2)
            Else
                Dim dtNow As Date = Now
                With dtNow
                    Dim Parms As String() = { _
                        Root.Root.ToString, _
                        "FileList", _
                        Root.Name, _
                        .Year.ToString & .Month.ToString("00") & .Day.ToString("00"), _
                        .Hour.ToString("00") & .Minute.ToString("00") & .Second.ToString("00")}
                    OutputFileName = String.Format("{0}{1}-{2}.{3}.{4}.csv", Parms)
                End With
            End If
            sw = New StreamWriter(OutputFileName)
            sw.WriteLine(String.Format("{0},{1},{2},{3},{4},{5}", Entry))
            Try
                fl.ListFiles(Root, sw)
                sw.Close()
            Catch ex As Exception
                sw.WriteLine(ex.ToString)
            End Try
            Return 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "FileList", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Function
End Class
