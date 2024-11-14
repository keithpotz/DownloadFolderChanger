' Download Folder Changer
' Copyright (c) 2024 Keith Pottratz
' Licensed under the MIT License. See LICENSE file in the project root for full license information.



Imports Microsoft.Win32
Imports System.Drawing

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim toolTip As New ToolTip()
        toolTip.SetToolTip(btnSelectFolder, "Select a folder to set as the default download location.")
        toolTip.SetToolTip(btnApply, "Apply the selected folder as the default download location.")
        toolTip.SetToolTip(btnReset, "Reset to the original default download folder.")
    End Sub

    Private Sub btnSelectFolder_Click(sender As Object, e As EventArgs) Handles btnSelectFolder.Click
        Using folderDialog As New FolderBrowserDialog()
            folderDialog.Description = "Select the new default download folder"

            ' Show dialog and check if the user selected a folder
            If folderDialog.ShowDialog() = DialogResult.OK Then
                ' Display the selected folder path in the TextBox
                txtFolderPath.Text = folderDialog.SelectedPath
                LogStatus("Selected folder: " & folderDialog.SelectedPath)
            End If
        End Using
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        Dim downloadsPath As String = txtFolderPath.Text

        ' Validate the path before proceeding
        If String.IsNullOrWhiteSpace(downloadsPath) Then
            LogStatus("Please select a valid folder.", False)
            Return
        End If

        Try
            Dim userShellPath As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            Dim shellPath As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            Dim downloadsKey As String = "{374DE290-123F-4565-9164-39C4925E467B}"

            ' Modify the registry for the default Downloads folder
            Dim userShellFolders As RegistryKey = Registry.CurrentUser.OpenSubKey(userShellPath, True)
            Dim shellFolders As RegistryKey = Registry.CurrentUser.OpenSubKey(shellPath, True)

            If userShellFolders IsNot Nothing AndAlso shellFolders IsNot Nothing Then
                LogStatus($"Setting {userShellPath}\{downloadsKey} to: {downloadsPath}", True)
                userShellFolders.SetValue(downloadsKey, downloadsPath)

                LogStatus($"Setting {shellPath}\{downloadsKey} to: {downloadsPath}", True)
                shellFolders.SetValue(downloadsKey, downloadsPath)

                LogStatus("Default download folder updated successfully.", True)
            Else
                LogStatus("Error: Could not open registry keys.", False)
            End If
        Catch ex As UnauthorizedAccessException
            LogStatus("Permission denied. Please run program as Administrator.", False)
        Catch ex As Exception
            LogStatus("Error: " & ex.Message, False)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Dim downloadsPath As String = txtFolderPath.Text

        If String.IsNullOrWhiteSpace(downloadsPath) Then
            LogStatus("Please select a valid folder.", False)
            Return
        End If

        If Not IO.Directory.Exists(downloadsPath) Then
            LogStatus("The selected Folder does not exist. Please chose a valid folder", False)
            Return

        End If
        Try
            Dim userShellPath As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            Dim shellPath As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            Dim downloadsKey As String = "{374DE290-123F-4565-9164-39C4925E467B}"

            ' Open registry keys for modifying
            Dim userShellFolders As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", True)
            Dim shellFolders As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", True)

            If userShellFolders IsNot Nothing AndAlso shellFolders IsNot Nothing Then
                ' Remove custom download folder settings to revert to default

                LogStatus($"Removing custom path from {userShellPath}\{downloadsKey} to reset to default.", True)
                userShellFolders.DeleteValue(downloadsKey, False)

                LogStatus($"Removing custom path from {shellPath}\{downloadsKey} to reset to default.", True)
                shellFolders.DeleteValue(downloadsKey, False)

                LogStatus("Default download folder reset to system default.", True)
            Else
                LogStatus("Error: Could not open registry keys.", False)
            End If
        Catch ex As Exception
            LogStatus("Error: " & ex.Message, False)
        End Try
    End Sub

    Private Sub LogStatus(message As String, Optional isSuccess As Boolean = True)
        If rtbStatus.Lines.Length > 100 Then
            rtbStatus.Text = String.Join(Environment.NewLine, rtbStatus.Lines.Skip(1))
        End If
        rtbStatus.AppendText($"{DateTime.Now}: {message}" & Environment.NewLine)
        rtbStatus.ScrollToCaret()

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If rtbStatus IsNot Nothing Then
            rtbStatus.Dispose()
        End If
        GC.Collect()
    End Sub
End Class
