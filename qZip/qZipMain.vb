Imports System.IO
Imports System.IO.Compression

Public Class qZipMain

    Private Shared FS As Char = Path.DirectorySeparatorChar
    Private Const DUPLICATE_PATH_SUFFIX = " ({0})"

    Private Sub Start() Handles Me.Shown

        '' Get the file and update the label.
        Dim targetFile = getTargetFile()
        Label1.Text = Label1.Text.Replace("{filename}", targetFile.Name)
        Label1.Refresh()

        getParentProcess()

        '' Get the temporary path and then unzip the folder.  Delete the temp path at the end regardless of success or failure.
        Dim tempFolder As New DirectoryInfo(targetFile.DirectoryName & FS & ".tmp-" & Guid.NewGuid().ToString())
        Try
            UnZip(targetFile, tempFolder)
        Finally
            If tempFolder.Exists Then tempFolder.Delete()
        End Try

        '' Exit the application.
        Me.Close()
    End Sub

    Function getTargetFile() As FileInfo
        '' Ensure valid filename was given.
        Dim targetFilename = Environment.GetCommandLineArgs().ElementAtOrDefault(1)
        While Not File.Exists(targetFilename)
            Dim ofd As New OpenFileDialog() With {
                .Title = "Select file to Unzip...",
                .Filter = "Zip Files (*.zip)|*.zip|All Files (*.*)|*.*"
            }
            If ofd.ShowDialog() = DialogResult.OK Then targetFilename = ofd.FileName
            ''Dim errorMessage = String.Format("Invalid filename specified: ""{0}"".", targetFilename)
            ''MessageBox.Show(errorMessage, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End While

        '' Return the FileInfo object.
        Return New FileInfo(targetFilename)
    End Function

    Sub UnZip(targetFile As FileInfo, tempFolder As DirectoryInfo)
        'tempFolder.Create()
        'tempFolder.Attributes = tempFolder.Attributes Or FileAttributes.Hidden
        ZipFile.ExtractToDirectory(targetFile.FullName, tempFolder.FullName)

        Dim destination As String
        Dim contents = tempFolder.GetFileSystemInfos()
        If (contents.Length = 0) Then
            MessageBox.Show("Zip file was empty.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf contents.Length > 1 Then
            destination = getUniquePathName(targetFile.Directory, getNameWithoutExt(targetFile))
            Directory.Move(tempFolder.FullName, destination)
        ElseIf contents.First.Attributes.HasFlag(FileAttributes.Directory) Then
            destination = getUniquePathName(targetFile.Directory, contents.First.Name)
            Directory.Move(contents.First.FullName, destination)
        Else
            destination = getUniquePathName(targetFile.Directory, getNameWithoutExt(contents.First), contents.First.Extension)
            File.Move(contents.First.FullName, destination)
        End If

        Process.Start("explorer.exe", String.Format("/select,""{0}""", destination))
        Dim parentProcess = getParentProcess()
        If parentProcess.ProcessName = "explorer" Then parentProcess.Kill()

        '' Remove the original zipped file.
        targetFile.Delete()
    End Sub

    Function getParentProcess() As Process
        Dim myId = Process.GetCurrentProcess().Id
        Dim query = String.Format("SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {0}", myId)
        Dim search = New System.Management.ManagementObjectSearcher("root\CIMV2", query)
        Dim results = search.[Get]().GetEnumerator()
        results.MoveNext()
        Dim queryObj = results.Current
        Dim parentId = CUInt(queryObj("ParentProcessId"))
        Return Process.GetProcessById(CInt(parentId))
    End Function

    Function getNameWithoutExt(file As FileSystemInfo) As String
        If (file.Attributes.HasFlag(FileAttributes.Directory)) Then Return file.Name
        Dim nameLength As Integer = file.Name.Length - file.Extension.Length
        Return file.Name.Substring(0, nameLength)
    End Function

    Function getUniquePathName(dirInfo As DirectoryInfo, name As String, Optional extension As String = "") As String
        Dim pathStart = dirInfo.FullName & FS & name
        Dim fullPath = pathStart & extension
        Dim i As Integer = 2 ' First copy with no suffix is the original copy, so the first copy to use a suffix will be "2"
        While File.Exists(fullPath) OrElse Directory.Exists(fullPath)
            fullPath = pathStart & String.Format(DUPLICATE_PATH_SUFFIX, i) & extension
            i += 1
        End While
        Return fullPath
    End Function


End Class
