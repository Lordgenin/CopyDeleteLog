Imports System.IO
Imports System.Windows.Forms

Public Class Form1
    Dim sourceLabel As New Label() With {.Text = "Source", .Location = New Point(10, 10)}
    Dim destinationLabel As New Label() With {.Text = "Destination", .Location = New Point(10, 70)}
    Dim deleteLabel As New Label() With {.Text = "Delete", .Location = New Point(10, 130)}

    Dim sourcePathTextBox As New TextBox() With {.Location = New Point(125, 10), .Width = 380, .Multiline = True, .Height = 50}
    Dim destinationPathTextBox As New TextBox() With {.Location = New Point(125, 70), .Width = 380, .Multiline = True, .Height = 50}
    Dim deletePathTextBox As New TextBox() With {.Location = New Point(125, 130), .Width = 380, .Multiline = True, .Height = 50}

    Dim copyButton As New Button() With {.Text = "Copy", .Location = New Point(10, 190)}
    Dim deleteButton As New Button() With {.Text = "Delete All", .Location = New Point(100, 190)}

    Dim logPathTextBox As New TextBox() With {.Location = New Point(125, 250), .Width = 380}
    Dim LogFilePath As New Button() With {.Text = "Select Log File Path", .Location = New Point(10, 250)}


    Public Sub New()
        Controls.Add(sourceLabel)
        Controls.Add(destinationLabel)
        Controls.Add(deleteLabel)
        Controls.Add(sourcePathTextBox)
        Controls.Add(destinationPathTextBox)
        Controls.Add(deletePathTextBox)
        Controls.Add(copyButton)
        Controls.Add(deleteButton)
        AddHandler copyButton.Click, AddressOf copyButton_Click
        AddHandler deleteButton.Click, AddressOf deleteButton_Click
        Controls.Add(logPathTextBox)
        Controls.Add(LogFilePath)
        AddHandler LogFilePath.Click, AddressOf LogFilePath_Click
    End Sub

    Private Sub LogFilePath_Click(sender As Object, e As EventArgs)
        Using sfd As New SaveFileDialog()
            sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            sfd.Title = "Select Log File Path"
            If sfd.ShowDialog() = DialogResult.OK Then
                logPathTextBox.Text = sfd.FileName
            End If
        End Using
    End Sub
    Private Sub copyButton_Click(sender As Object, e As EventArgs)
        Try
            Dim sourcePath As String = sourcePathTextBox.Text
            Dim destinationPath As String = destinationPathTextBox.Text

            ' Validate source and destination paths
            If String.IsNullOrWhiteSpace(sourcePath) OrElse String.IsNullOrWhiteSpace(destinationPath) Then
                MessageBox.Show("Source and destination paths cannot be empty!")
                Return
            End If

            If Not Directory.Exists(sourcePath) Then
                MessageBox.Show("Source directory does not exist!")
                Return
            End If

            If File.Exists(destinationPath) Then
                MessageBox.Show("Destination path points to an existing file!")
                Return
            End If

            ' Copy directory to destination
            LoadLoggedData() ' Load the logged directories and files into the HashSets
            CopyDirectory(sourcePath, destinationPath)
            MessageBox.Show("Directory copied successfully!")
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub deleteButton_Click(sender As Object, e As EventArgs)
        Try
            Dim deletePath As String = deletePathTextBox.Text

            ' Validate delete path
            If String.IsNullOrWhiteSpace(deletePath) Then
                MessageBox.Show("Delete path cannot be empty!")
                Return
            End If

            If Not Directory.Exists(deletePath) Then
                MessageBox.Show("Delete directory does not exist!")
                Return
            End If

            ' Delete directory and all its content
            Directory.Delete(deletePath, True)

            MessageBox.Show("All files and subdirectories in delete directory have been deleted!")
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub


    Private Sub LogDirectory(directoryPath As String)
        If String.IsNullOrEmpty(logPathTextBox.Text) Then
            MessageBox.Show("Please select a log file path first.")
            Return
        End If

        Using sw As StreamWriter = File.AppendText(logPathTextBox.Text)
            sw.WriteLine("DIR: " & directoryPath)
        End Using
    End Sub


    Private Sub CopyDirectory(sourceDir As String, destDir As String)
        Dim directoriesToCopy As New Stack(Of Tuple(Of String, String))()
        directoriesToCopy.Push(New Tuple(Of String, String)(sourceDir, destDir))

        ' If the loaded log is empty, assume everything needs to be copied.
        Dim isFirstCopy As Boolean = (loggedDirectories.Count = 0)

        While directoriesToCopy.Count > 0
            Dim current = directoriesToCopy.Pop()
            Dim currentSourceDir = current.Item1
            Dim currentDestDir = current.Item2



            ' Check if the current source directory hasn't been logged or it's the first copy.
            If isFirstCopy OrElse Not loggedDirectories.Contains(currentSourceDir) Then
                Directory.CreateDirectory(currentDestDir)

                For Each file In Directory.GetFiles(currentSourceDir)
                    Dim destFilePath = Path.Combine(currentDestDir, Path.GetFileName(file))

                    If Not System.IO.File.Exists(destFilePath) AndAlso Not IsFileLogged(file) Then
                        System.IO.File.Copy(file, destFilePath, True) ' The third parameter (True) allows overwriting if the destination file already exists
                        LogFile(file)
                    End If
                Next

                For Each subDir In Directory.GetDirectories(currentSourceDir)
                    Dim destSubDir = Path.Combine(currentDestDir, Path.GetFileName(subDir))
                    directoriesToCopy.Push(New Tuple(Of String, String)(subDir, destSubDir))
                Next

                LogDirectory(currentSourceDir)
            End If
        End While
    End Sub





    Private Function IsFileLogged(filePath As String) As Boolean
        Return loggedFiles.Contains(filePath)
    End Function

    Private Sub LogFile(filePath As String)
        If String.IsNullOrEmpty(logPathTextBox.Text) Then
            MessageBox.Show("Please select a log file path first.")
            Return
        End If

        Using sw As StreamWriter = File.AppendText(logPathTextBox.Text)
            sw.WriteLine("FILE: " & filePath)
        End Using
    End Sub


    Private loggedDirectories As HashSet(Of String) = New HashSet(Of String)()
    Private loggedFiles As HashSet(Of String) = New HashSet(Of String)()

    Private Sub LoadLoggedData()
        If File.Exists(logPathTextBox.Text) Then
            For Each line In File.ReadLines(logPathTextBox.Text)
                If line.StartsWith("DIR: ") Then
                    loggedDirectories.Add(line.Substring(5))
                ElseIf line.StartsWith("FILE: ") Then
                    loggedFiles.Add(line.Substring(6))
                End If
            Next
        End If
    End Sub
End Class
