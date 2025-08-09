Imports System.Data
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Reporting.WinForms
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Public Class frmVerisys
    Private connDBverisys As SqlConnection
    Private cmdVerisys As SqlCommand
    Private readerVerisys As SqlDataReader
    Private cmdLoc As SqlCommand
    Private readerLoc As SqlDataReader
    Public Sub openConn()
        connDBverisys = New SqlConnection("Server=" + ServerName + ";Database=DB2023VERS;Trusted_Connection=True;")
        'connDBverisys.Open()
        If connDBverisys.State = ConnectionState.Closed Then
            connDBverisys.Open()
        End If
    End Sub
    Public Sub MoveVerisys(ByVal sourceFilePath As String, ByVal destinationFolderPath As String)
        'Dim sourceFilePath As String = "C:\SourceFolder\file.txt"
        'Dim destinationFolderPath As String = "C:\DestinationFolder"

        Try
            ' Check if the source file exists before attempting to move it
            If File.Exists(sourceFilePath) Then
                ' Combine the destination folder path with the original file name to get the full destination path
                Dim destinationFilePath As String = Path.Combine(destinationFolderPath, Path.GetFileName(sourceFilePath))

                ' Perform the move operation
                File.Move(sourceFilePath, destinationFilePath)

                MessageBox.Show("File moved successfully!")
            Else
                MessageBox.Show("Source file does not exist.")
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub
    Public Sub insertLocation(ByVal location As String)
        Dim insertQuery As String = "insert into TB2023LOC (Location) VALUE(" + location + ")"
        Dim recordExists As Boolean = CheckRecordExists("TB2023LOC", "Location", location)

        If recordExists Then
            ' Record exists
            MsgBox("Record already exists.")
        Else
            ' Record does not exist

            cmdLoc = New SqlCommand(insertQuery, connDBverisys)
            openConn()
            Try
                cmdLoc.ExecuteNonQuery()
                connDBverisys.Close()
            Catch ex As Exception
                MsgBox("Error during inserting location" + vbCrLf + Err.Description, MsgBoxStyle.Information)
                connDBverisys.Close()
            End Try
        End If


    End Sub

    Public Sub UpdateLocation()

    End Sub
    Public Sub deleteLocation()

    End Sub
    Public Sub insertVerisys(ByVal ID As String, ByVal CNIC As String)
        Dim insertQuery As String = "insert into TB2023VERS (ID,CNIC) VALUE(" + ID + "," + CNIC + ")"
        Dim recordExists As Boolean = CheckRecordExists("TB2023VERS", "CNIC", CNIC)

        If recordExists = True Then
            ' Record exists
            MsgBox("Record already exists.")
        Else
            ' Record does not exist

            cmdLoc = New SqlCommand(insertQuery, connDBverisys)
            openConn()
            Try
                cmdLoc.ExecuteNonQuery()
                connDBverisys.Close()
            Catch ex As Exception
                MsgBox("Error during inserting CNIC" + vbCrLf + Err.Description, MsgBoxStyle.Information)
                connDBverisys.Close()
            End Try
        End If
    End Sub
    Public Sub updateVerisys()

    End Sub
    Public Sub deleteVerisys()

    End Sub
    Public Function CheckRecordExists(tableName As String, columnName As String, value As Object) As Boolean
        Dim recordExists As Boolean = False

        ' Create the SQL connection
        openConn()

        ' Create the SQL command
        ' Dim command1 As SqlCommand
        Dim commandText As String = "SELECT COUNT(*) FROM " + tableName + " WHERE " + columnName + " = @Value"
        Using command As New SqlCommand(commandText, connDBverisys)
            ' Add the parameter
            command.Parameters.AddWithValue("@Value", value)

            ' Execute the command and get the count
            Dim count As Integer = CInt(command.ExecuteScalar())
            connDBverisys.Close()
            command.Dispose()
            ' Check if the count is greater than 0
            If count > 0 Then
                recordExists = True
            End If
        End Using


        Return recordExists
    End Function
    Public AllCommonFileNames() As String
    Public OnlyFilesPath As String
    Dim PathAndFileName As String = Nothing
    Dim OnlyFileName As String = Nothing
    Dim fileNames() As String
    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Photos|*.JPG;*.PNG"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
            lbSelectedFiles.Text = NumberOfFiles
            txtVerisysPath.Text = Path.GetDirectoryName(OpenFileDialog1.FileName)
            pbVerisys.Maximum = NumberOfFiles
            pbVerisys.Minimum = 0
            pbVerisys.Value = 0
            pbVerisys.Refresh()
            
            AllCommonFileNames = OpenFileDialog1.FileNames()
            Dim directoryPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'fileNames = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileNames())
            OnlyFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            ' OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            OnlyFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            'If FileExtention = ".xlsx" Then
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            'Else
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            'End If
            'txt_CDR_CommonLinks_Path.Text = PathAndFileName
            'frmXLFormat.NoOfFiles = AllCommonFileNames
            'frmXLFormat.ShowDialog()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub frmVerisys_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        frm_Spy_Tech.Enabled = False
        Me.CenterToScreen()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        frm_Spy_Tech.Enabled = True
        Me.Close()
    End Sub

    Private Sub cbSameAddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSameAddress.CheckedChanged
        If cbSameAddress.Checked = True Then
            txtTargetPath.Text = txtVerisysPath.Text
            btnTargetFolder.Enabled = False
        Else
            btnTargetFolder.Enabled = True
            txtTargetPath.Text = ""
        End If
    End Sub

    Private Sub btnTargetFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTargetFolder.Click
        Dim folderbrowser As New FolderBrowserDialog
        If folderbrowser.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' Dim AllFileNames() As String = OpenFileDialog1.FileNames()

            txtTargetPath.Text = folderbrowser.SelectedPath


            'AllCommonFileNames = OpenFileDialog1.FileNames()
            'Dim directoryPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
            ''fileNames = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileNames())
            'OnlyFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            '' OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            'FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            'OnlyFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            ''If FileExtention = ".xlsx" Then
            ''OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            ''Else
            ''OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            ''End If
            ''txt_CDR_CommonLinks_Path.Text = PathAndFileName
            ''frmXLFormat.NoOfFiles = AllCommonFileNames
            ''frmXLFormat.ShowDialog()
        Else
            Exit Sub
        End If
    End Sub
End Class