Imports System.Data
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
'Imports Microsoft.Reporting.WinForms
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.SqlClient
'Imports Microsoft.Office.Interop
Imports System.IO
Public Class frmMergeExcels


    Private Sub frmMergeExcels_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterParent
        Call RefreshForm()
        txtFolderName.Text = "Enter FIR Year Police Station"
        txtFolderName.ForeColor = Color.Gray
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Dim AllMobilinkFileNames() As String
    Dim MobilinkFileName As String
    Dim MobilinkFilesPath As String
    Dim MobilinkBTSFileName As String
    Private Sub btnMobilink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMobilink.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim FileExtention As String = Nothing
        OpenFileDialog1.Title = "Select & Open Mobilink files"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Excel|*.xls;*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
            'lbSelectedFiles.Text = NumberOfFiles
            'txtVerisysPath.Text = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'pbVerisys.Maximum = NumberOfFiles
            'pbVerisys.Minimum = 0
            'pbVerisys.Value = 0
            'pbVerisys.Refresh()
            lbMobilink.Text = NumberOfFiles & " File(s) Selected"
            AllMobilinkFileNames = OpenFileDialog1.FileNames()
            'Dim directoryPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'fileNames = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileNames())
            MobilinkFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            ' OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            MobilinkFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

            'If FileExtention = ".xlsx" Then
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            'Else
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            'End If
            'txt_CDR_CommonLinks_Path.Text = PathAndFileName
            'frmXLFormat.NoOfFiles = AllCommonFileNames
            'frmXLFormat.ShowDialog()
            isMobilink = True
            MobilinkBTSFileName = "Mobilink BTS"
        Else
            Exit Sub
        End If
    End Sub
    Dim AllUfoneFileNames() As String
    Dim UfoneFileName As String
    Dim UfoneFilesPath As String
    Dim UfoneBTSFileName As String
    Private Sub btnUfone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUfone.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Title = "Select & Open Ufone files"
        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Text Excel|*.txt;*.xls;*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
            'lbSelectedFiles.Text = NumberOfFiles
            'txtVerisysPath.Text = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'pbVerisys.Maximum = NumberOfFiles
            'pbVerisys.Minimum = 0
            'pbVerisys.Value = 0
            'pbVerisys.Refresh()
            lbUfone.Text = NumberOfFiles & " File(s) Selected"
            AllUfoneFileNames = OpenFileDialog1.FileNames()
            'Dim directoryPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'fileNames = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileNames())
            UfoneFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            ' OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            UfoneFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

            'If FileExtention = ".xlsx" Then
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            'Else
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            'End If
            'txt_CDR_CommonLinks_Path.Text = PathAndFileName
            'frmXLFormat.NoOfFiles = AllCommonFileNames
            'frmXLFormat.ShowDialog()
            isUfone = True
            UfoneBTSFileName = "Ufone BTS"
        Else
            Exit Sub
        End If
    End Sub
    Dim AllZongFileNames() As String
    Dim ZongFileName As String
    Dim ZongFilesPath As String
    Dim ZongBTSFileName As String
    Private Sub btnZong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZong.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Title = "Select & Open Zong files"
        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Excel|*.xls;*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
            'lbSelectedFiles.Text = NumberOfFiles
            'txtVerisysPath.Text = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'pbVerisys.Maximum = NumberOfFiles
            'pbVerisys.Minimum = 0
            'pbVerisys.Value = 0
            'pbVerisys.Refresh()
            lbZong.Text = NumberOfFiles & " File(s) Selected"
            AllZongFileNames = OpenFileDialog1.FileNames()
            'Dim directoryPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'fileNames = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileNames())
            ZongFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            ' OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            ZongFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

            'If FileExtention = ".xlsx" Then
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            'Else
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            'End If
            'txt_CDR_CommonLinks_Path.Text = PathAndFileName
            'frmXLFormat.NoOfFiles = AllCommonFileNames
            'frmXLFormat.ShowDialog()
            isZong = True
            ZongBTSFileName = "Zong BTS"
        Else
            Exit Sub
        End If
    End Sub
    Dim AllTeleNorFileNames() As String
    Dim TeleNorFileName As String
    Dim TeleNorFilesPath As String
    Dim TeleNorBTSFileName As String
    Private Sub btnTeleNor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTeleNor.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Title = "Select & Open TeleNor files"
        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Comma Separated Values|*.csv"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
           
            lbTeleNor.Text = NumberOfFiles & " File(s) Selected"
            AllTeleNorFileNames = OpenFileDialog1.FileNames()
            'Dim directoryPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
            'fileNames = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileNames())
            TeleNorFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            ' OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            TeleNorFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

            'If FileExtention = ".xlsx" Then
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            'Else
            'OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            'End If
            'txt_CDR_CommonLinks_Path.Text = PathAndFileName
            'frmXLFormat.NoOfFiles = AllCommonFileNames
            'frmXLFormat.ShowDialog()
            isTeleNor = True
            TeleNorBTSFileName = "TeleNor BTS"
        Else
            Exit Sub
        End If
    End Sub
    Public newXlApp As Excel.Application
    Public newXlWorkbooks As Excel.Workbooks
    Public newXlWorkbook As Excel.Workbook
    Public newXlWorkSheet As Excel.Worksheet
    Public xlSheets As Excel.Worksheets
    'where to insert current row in the target xl file
    Public xlRowNumber As Integer
    'Number of records in current xl file
    Public NumberRecs As Integer
    Private MergeXLApp As Excel.Application
    Private MergeXLWorkbooks As Excel.Workbooks
    Private MergeXLWorkbook As Excel.Workbook
    Private MergeXLWorkSheet As Excel.Worksheet
    Private Sub CreateNewExcelFile()
        newXlApp = New Excel.Application
        newXlWorkbook = newXlApp.Workbooks.Add()
        Dim misValue As Object = System.Reflection.Missing.Value
        'newXlWorkSheet = "bts"
        newXlWorkSheet = newXlWorkbook.Sheets("Sheet1")
        newXlWorkSheet.Name = "bts"
        newXlWorkSheet = newXlWorkbook.Sheets("bts")
        xlRowNumber = 1
        NumberRecs = 0
        For i As Integer = 1 To 10
            If i = 1 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "a"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 2 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "b"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 3 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "Time"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@" '"hh:MM:ss AM/PM"
            ElseIf i = 4 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "Date"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@" '"dd/mm/yyyy"
            ElseIf i = 5 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "Call Type"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 6 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "Duration"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 7 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "IMEI"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 8 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "IMSI"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 9 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "Cell ID"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
            ElseIf i = 10 Then
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).Value = "Site"
                newXlWorkSheet.Range(newXlWorkSheet.Cells(1, i), newXlWorkSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"

            End If
        Next
        newXlApp.DisplayAlerts = False
        'Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CompRpt_CallSammary")
        Dim excelFilePath As String = "C:\Users\sajjad\Desktop\Sample\BTS\Sample6.xlsx"
        Try
            newXlWorkbook.SaveAs(excelFilePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        Catch ex As Exception
            MsgBox("Error in saving Comprehaensive CDR Report " & Err.Description, MsgBoxStyle.Information)
        End Try
        newXlWorkbook.Close()
        'newXlWorkbooks.Close()
        newXlApp.Quit()
        releaseObject(newXlWorkSheet)
        'releaseObject(xlDataSheet)
        releaseObject(newXlWorkbook)
        'releaseObject(newXlWorkbooks)
        releaseObject(newXlApp)

        'for InsertRec method

        'MergeXLApp = New Excel.Application
        'MergeXLWorkbooks = MergeXLApp.Workbooks
        'MergeXLWorkbook = MergeXLWorkbooks.Open(excelFilePath)
        ''newXlWorkSheet = "bts"
        'MergeXLWorkSheet = MergeXLWorkbook.Sheets("bts")


        
        'SheetName = "bts$"
    End Sub
    Private Sub InsertRec(ByVal rownumber As Integer, ByVal aParty As String, ByVal bParty As String, ByVal calltime As Date, ByVal calldate As Date, ByVal CallType As String, ByVal duration As String, Optional ByVal imei As String = "", Optional ByVal imsi As String = "", Optional ByVal cellid As String = "", Optional ByVal cite As String = "")


        MergeXLWorkSheet.Cells(rownumber, 1) = aParty
        MergeXLWorkSheet.Cells(rownumber, 2) = bParty
        MergeXLWorkSheet.Cells(rownumber, 3) = calltime
        MergeXLWorkSheet.Cells(rownumber, 4) = calldate
        MergeXLWorkSheet.Cells(rownumber, 5) = CallType
        MergeXLWorkSheet.Cells(rownumber, 6) = duration
        MergeXLWorkSheet.Cells(rownumber, 7) = imei
        MergeXLWorkSheet.Cells(rownumber, 8) = imsi
        MergeXLWorkSheet.Cells(rownumber, 9) = cellid
        MergeXLWorkSheet.Cells(rownumber, 10) = cite
        lbRecCount.Text = rownumber
        lbRecCount.Refresh()
    End Sub
    Private Sub RetrieveRecs()
        'Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        'Dim o As New OleDb.OleDbConnection(ConnectionString)
        'o.Open()
        'Dim IsbpartyAdded As Boolean = False
        'Dim NextNumber As Long = Nothing
        'Dim NextNumbertxt As String = Nothing
        'Dim ocmd1 As New OleDb.OleDbCommand(queryString_b, o)
        ''ocmd1.ExecuteNonQuery()
        '' Dim oreader As OleDb.OleDbDataReader
        'Dim oreader_b As OleDb.OleDbDataReader
        'oreader_b = ocmd1.ExecuteReader()
        'totalfields = oreader_b.FieldCount
        'For j As Integer = 1 To totalfields
        '    newXlWorkSheet.Cells(1, j) = oreader_b.GetName(j - 1).ToString
        'Next
    End Sub
    Private Sub Mobilink_to_BTS(ByVal CellID_File_Path As String, ByVal SheetName As String)
        Dim oledbConn As OleDb.OleDbConnection
        Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & CellID_File_Path & ";extended properties='excel 12.0 xml; hdr=yes'"
        oledbConn = New OleDb.OleDbConnection(oledbconnstr)
        oledbConn.Open()
        Dim SelectCommand As New OleDb.OleDbCommand
        Dim oledbReader As OleDbDataReader
        SelectCommand.CommandText = "Select * from [" & SheetName & "$" & "]"
        SelectCommand.Connection = oledbConn
        oledbReader = SelectCommand.ExecuteReader()
        'For j As Integer = 0 To TotalNumberOfCDRs - 1
        '    OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
        'Next
        Dim insertquery As String
        CreatSQLTable("bts")
        While oledbReader.Read()
            xlRowNumber = xlRowNumber + 1
            'InsertMobilinkRecs("C:\Users\sajjad\Desktop\Sample\BTS\Sample6.xlsx", "bts", oledbReader("A-Party"), oledbReader("b-Party"), DateTime.Parse(oledbReader("Date & Time"))) ', oledbReader("Date & Time"))
            'InsertRec(xlRowNumber, oledbReader("A-Party"), oledbReader("B-Party"), DateTime.Parse(oledbReader("Date & Time")), DateTime.Parse(oledbReader("Date & Time")), oledbReader("Call Type"), oledbReader("Duration"), oledbReader("IMEI"), "", "", "")
            'insertquery = "INSERT INTO [" & SheetName & "$" & "] (a) VALUES ('" & oledbReader("A-Party") & "')"
            '& "' ,'oledbReader("B-Party"), DateTime.Parse(oledbReader("Date & Time")), DateTime.Parse(oledbReader("Date & Time")), oledbReader("Call Type"), oledbReader("Duration"), oledbReader("IMEI"))"
        End While
        oledbConn.Close()
        oledbReader.Close()
        SelectCommand.Dispose()
    End Sub
    Private Sub CreatSQLTable(ByVal TempTableName As String)
        Dim TargetConnection As SqlConnection = New SqlConnection("Server=" + ServerName + ";Database=tempdb;Trusted_Connection=True;")
        TargetConnection.Open()
        Dim OthersCommand As SqlCommand
        Dim CreateTbCommand As SqlCommand
        Dim queryString1, CreateTablesQuery As String
        Try
            queryString1 = "IF OBJECT_ID('dbo." & TempTableName & "') IS NOT NULL DROP TABLE " & TempTableName & ""
            CreateTbCommand = New SqlCommand(queryString1, OthersConnection)
            CreateTbCommand.ExecuteNonQuery()
        Catch ex03 As Exception
        End Try

        CreateTablesQuery = "CREATE TABLE [" & TempTableName & "] (PhoneNumber nvarchar(255) null)"
        OthersCommand = New SqlCommand(CreateTablesQuery, OthersConnection)
        Try
            OthersCommand.ExecuteNonQuery()
        Catch ex02 As Exception
            MsgBox("creating table", MsgBoxStyle.OkOnly)
        End Try
        TargetConnection.Close()
        OthersCommand.Dispose()
        CreateTbCommand.Dispose()
        queryString1 = Nothing
        CreateTablesQuery = Nothing
        System.GC.Collect()
    End Sub
    Private Sub InsertMobilinkRecs(ByVal MergingFilePath As String, ByVal SheetName As String, ByVal a As String, ByVal b As String, ByVal CallTime As Date) ', ByVal CallDate As String)
        Dim oledbConn As OleDb.OleDbConnection
        Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & MergingFilePath & ";extended properties='excel 12.0 xml; hdr=yes'"
        oledbConn = New OleDb.OleDbConnection(oledbconnstr)
        oledbConn.Open()
        Dim SelectCommand As OleDb.OleDbCommand
        'Dim oledbReader As OleDbDataReader
        'Dim formattedDate As String = CallTime.ToString("hh:mm:ss tt")
        Dim insertQuery As String = "INSERT INTO [" & SheetName & "$" & "] (a,b,Time) VALUES ('" & a & "' , '" & b & "' , " & CallTime & ")"
        ' , '{" & CallDate & "}')"
        'SelectCommand= New OleDbCommand(insertQuery, oledbConn))
        'SelectCommand.Parameters.AddWithValue("", )
        SelectCommand.ExecuteNonQuery()
        oledbConn.Close()
        'oledbReader.Close()
        SelectCommand.Dispose()
        'oledbReader = SelectCommand.ExecuteReader()
        'For j As Integer = 0 To TotalNumberOfCDRs - 1
        '    OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
        'Next
        'Dim insertquery As String
        'insertquery = "INSERT INTO [" & SheetName & "$" & "] (a) VALUES ('" & oledbReader("A-Party") & "')"

        'While oledbReader.Read()
        '    xlRowNumber = xlRowNumber + 1
        '    'InsertRec(xlRowNumber, oledbReader("A-Party"), oledbReader("B-Party"), DateTime.Parse(oledbReader("Date & Time")), DateTime.Parse(oledbReader("Date & Time")), oledbReader("Call Type"), oledbReader("Duration"), oledbReader("IMEI"), "", "", "")

        '    '& "' ,'oledbReader("B-Party"), DateTime.Parse(oledbReader("Date & Time")), DateTime.Parse(oledbReader("Date & Time")), oledbReader("Call Type"), oledbReader("Duration"), oledbReader("IMEI"))"
        'End While
    End Sub
    Private Sub MergingXLFiles(ByVal XLMergingFiles() As String, ByVal NoOfMergingFiles As Integer)
        xlRowNumber = 0
        Dim FinishedFiles As String = 0
        'csvFilePath = Path.GetDirectoryName(XLMergingFiles(0)) & "\bts.csv"
        'csvFilePath = Path.GetDirectoryName(XLMergingFiles(0)) & "\" & Company & " BTS.csv"
        Dim regkey As Object = My.Computer.Registry.CurrentUser.GetValue("MegerPath")
        'csvFilePath = My.Computer.Registry.CurrentUser.GetValue("MegerPath") + "\" + txtFolderName.Text & "\" & Company & " BTS.csv"
        If Directory.GetDirectoryRoot(regkey) = regkey Then
            csvFilePath = My.Computer.Registry.CurrentUser.GetValue("MegerPath") + txtFolderName.Text & "\" & Company & " BTS.csv"
        Else
            csvFilePath = My.Computer.Registry.CurrentUser.GetValue("MegerPath") + "\" + txtFolderName.Text & "\" & Company & " BTS.csv"

        End If
        'If csvFilePath Is Nothing Then
        '    MsgBox("Please set Root Drive or Folder", MsgBoxStyle.OkOnly)
        '    Exit Sub
        'End If
       
        outFile = My.Computer.FileSystem.OpenTextFileWriter(csvFilePath, False)
        outFile.WriteLine("a" & "," & "b" & "," & "Time" & "," & "Date" & "," & "Call Type" & "," & "Duration" & "," & "IMEI" & "," & "IMSI" & "," & "Cell ID" & "," & "Site")
        
        For i As Integer = 0 To NoOfMergingFiles - 1
            'Mobilink_to_BTS(XLMergingFiles(i), GetSheetName(XLMergingFiles(i)))
            If Company = "Mobilink" Then
                lbMobilink.Text = XLMergingFiles(i)
                lbMobilink.Refresh()
                FinishedFiles = FinishedFiles + 1
                lbMobilinkFileCount.Text = FinishedFiles & " of " & NoOfMergingFiles
                lbMobilinkFileCount.Refresh()
                CreateCSVMobilink(XLMergingFiles(i), GetSheetName(XLMergingFiles(i)))
            End If
            If Company = "Ufone" Then
                lbUfone.Text = XLMergingFiles(i)
                lbUfone.Refresh()
                FinishedFiles = FinishedFiles + 1
                lbUfoneFileCount.Text = FinishedFiles & " of " & NoOfMergingFiles
                lbUfoneFileCount.Refresh()
                If Path.GetExtension(XLMergingFiles(i)) = ".txt" Then
                    CreateCsvUfone(XLMergingFiles(i))
                ElseIf Path.GetExtension(XLMergingFiles(i)) = ".xlsx" Or Path.GetExtension(XLMergingFiles(i)) = ".xls" Then
                    CreateCsvUfoneXLS(XLMergingFiles(i), GetSheetName(XLMergingFiles(i)))
                End If
            End If
            If Company = "Zong" Then
                lbZong.Text = XLMergingFiles(i)
                lbZong.Refresh()
                FinishedFiles = FinishedFiles + 1
                lbZongFileCount.Text = FinishedFiles & " of " & NoOfMergingFiles
                lbZongFileCount.Refresh()
                CreateCSVZong(XLMergingFiles(i), GetSheetName(XLMergingFiles(i)))
            End If
            If Company = "TeleNor" Then
                lbTeleNor.Text = XLMergingFiles(i)
                lbTeleNor.Refresh()
                FinishedFiles = FinishedFiles + 1
                lbTeleNorFileCount.Text = FinishedFiles & " of " & NoOfMergingFiles
                lbTeleNorFileCount.Refresh()
                CreateCsvTeleNor(XLMergingFiles(i))
            End If

        Next
        outFile.Close()
        Call labelsText(Company, "Finalizing ............. Step1")
        csvToExcel(csvFilePath)
    End Sub
    Function IsSheetEmpty(ByVal sheet As Excel.Worksheet) As Boolean
        ' Check if the sheet has any used cells
        Return sheet.UsedRange Is Nothing OrElse sheet.UsedRange.Cells.Count = 0
    End Function
    Function GetSheetName(ByVal xlFileAddress As String) As String
        ' Specify the path to your Excel file
        Dim filePath As String = xlFileAddress
        Dim sheetName As String
        ' Create an Excel application object
        Dim excelApp As New Excel.Application()

        ' Open the workbook
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(filePath)

        ' Iterate through all sheets in the workbook
        For Each sheet As Excel.Worksheet In workbook.Sheets
            ' Check if the sheet is not empty
            If Not IsSheetEmpty(sheet) Then
                sheetName = sheet.Name
                Exit For
            End If
        Next

        ' Close and release resources
        workbook.Close()
        excelApp.Quit()
        releaseObject(workbook)
        releaseObject(excelApp)
        Return sheetName
    End Function
    Private isMobilink As Boolean = False
    Private isUfone As Boolean = False
    Private isZong As Boolean = False
    Private isTeleNor As Boolean = False
    Function CreatFolder(ByVal FolderPath As String)
        If Not System.IO.Directory.Exists(FolderPath) Then
            System.IO.Directory.CreateDirectory(FolderPath)
        End If
    End Function
    Private Sub btnMerge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMerge.Click
        'Call CreateNewExcelFile()
        Dim regKey As Object = My.Computer.Registry.CurrentUser.GetValue("MegerPath")

        If regKey Is Nothing Then
            MsgBox("Please set Root Drive or Folder", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If txtFolderName.Text = "" Then
            MsgBox("Please give the Folder name in Term of FIR Year and Police Station name", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If System.IO.Directory.Exists(My.Computer.Registry.CurrentUser.GetValue("MegerPath") + "\" + txtFolderName.Text) Then
            Dim Ask As MsgBoxResult = MsgBox("The folder name already exists" + vbCrLf + "Would you like to continue....", MsgBoxStyle.YesNo)
            If Ask = MsgBoxResult.No Then
                txtFolderName.Focus()
                Exit Sub
            End If
        End If
        btnMerge.Enabled = False
        btnRefresh.Enabled = False
        If Directory.GetDirectoryRoot(regKey) = regKey Then
            Call CreatFolder(regKey + txtFolderName.Text)
        Else
            Call CreatFolder(regKey + "\" + txtFolderName.Text)
        End If

        If isMobilink = True Then
            prbMbilink.Maximum = AllMobilinkFileNames.Length + 3
            prbMbilink.Minimum = 0
            prbMbilink.Value = 0
            prbMbilink.Visible = True
            prbMbilink.Refresh()
            lbRecCount.Text = "0"
            lbRecCount.Visible = True
            Company = "Mobilink"
            lbMobilinkFileCount.Text = ""
            lbMobilinkFileCount.Visible = True
            Call MergingXLFiles(AllMobilinkFileNames, AllMobilinkFileNames.Length)
        End If
        If isUfone = True Then

            prbUfone.Maximum = AllUfoneFileNames.Length + 3
            prbUfone.Minimum = 0
            prbUfone.Value = 0
            prbUfone.Visible = True
            prbUfone.Refresh()
            lbUfoneCount.Text = "0"
            lbUfoneCount.Visible = True
            Company = "Ufone"
            lbUfoneFileCount.Text = ""
            lbUfoneFileCount.Visible = True
            Call MergingXLFiles(AllUfoneFileNames, AllUfoneFileNames.Length)
        End If
        If isZong = True Then
            prbZong.Maximum = AllZongFileNames.Length + 3
            prbZong.Minimum = 0
            prbZong.Value = 0
            prbZong.Visible = True
            prbZong.Refresh()
            lbZongCount.Text = "0"
            lbZongCount.Visible = True
            Company = "Zong"
            lbZongFileCount.Text = ""
            lbZongFileCount.Visible = True
            Call MergingXLFiles(AllZongFileNames, AllZongFileNames.Length)
        End If
        If isTeleNor = True Then
            prbTeleNor.Maximum = AllTeleNorFileNames.Length + 3
            prbTeleNor.Minimum = 0
            prbTeleNor.Value = 0
            prbTeleNor.Visible = True
            prbTeleNor.Refresh()
            lbTeleNorCount.Text = "0"
            lbTeleNorCount.Visible = True
            Company = "TeleNor"
            lbTeleNorFileCount.Text = ""
            lbTeleNorFileCount.Visible = True
            lbZongFileCount.Visible = True
            Call MergingXLFiles(AllTeleNorFileNames, AllTeleNorFileNames.Length)
        End If
        MsgBox("BTS(s) have been created:", MsgBoxStyle.OkOnly)
        btnMerge.Enabled = False
        btnRefresh.Enabled = True
        'CreateCSV(AllMobilinkFileNames, AllMobilinkFileNames.Length)
    End Sub
    Dim csvFilePath As String '= "C:\Users\sajjad\Desktop\Test.csv" 'Path to create or existing file
    Dim outFile As IO.StreamWriter '= My.Computer.FileSystem.OpenTextFileWriter(csvFilePath, False)
    Private Sub CreateCSVMobilink(ByVal CellID_File_Path As String, ByVal SheetName As String)
        Dim oledbConn As OleDb.OleDbConnection
        Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & CellID_File_Path & ";extended properties='excel 12.0 xml; hdr=yes'"
        oledbConn = New OleDb.OleDbConnection(oledbconnstr)
        oledbConn.Open()
        Dim SelectCommand As New OleDb.OleDbCommand
        Dim oledbReader As OleDbDataReader
        SelectCommand.CommandText = "Select * from [" & SheetName & "$" & "]"
        SelectCommand.Connection = oledbConn
        Dim MobilinkFileName As String = Path.GetFileNameWithoutExtension(CellID_File_Path)
        Try
            oledbReader = SelectCommand.ExecuteReader()
            While oledbReader.Read()
                xlRowNumber = xlRowNumber + 1
                outFile.WriteLine(oledbReader("A-Party") & "," & oledbReader("B-Party") & "," & DateToDouble(oledbReader("Date & Time")) & "," & DateToDouble(oledbReader("Date & Time")) & "," & oledbReader("Call Type") & "," & oledbReader("Duration") & "," & oledbReader("IMEI") & "," & "" & "," & MobilinkFileName)
                lbRecCount.Text = xlRowNumber
                lbRecCount.Refresh()
            End While
        Catch ex As Exception

        End Try
        oledbConn.Close()
        oledbReader.Close()
        SelectCommand.Dispose()
        prbMbilink.Value = prbMbilink.Value + 1
        prbMbilink.Refresh()
    End Sub
    Private Sub CreateCSVZong(ByVal CellID_File_Path As String, ByVal SheetName As String)
        Dim oledbConn As OleDb.OleDbConnection
        Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & CellID_File_Path & ";extended properties='excel 12.0 xml; hdr=yes'"
        oledbConn = New OleDb.OleDbConnection(oledbconnstr)
        oledbConn.Open()
        Dim SelectCommand As New OleDb.OleDbCommand
        Dim oledbReader As OleDbDataReader
        SelectCommand.CommandText = "Select * from [" & SheetName & "$" & "]"
        SelectCommand.Connection = oledbConn
        Try

        
        oledbReader = SelectCommand.ExecuteReader()
        Dim CallType As String
        Dim A_Party As String
        Dim B_Party As String
            Dim IMEI As String
            Dim IMSI As String = " "
            While oledbReader.Read()
                xlRowNumber = xlRowNumber + 1
                If oledbReader("DIR_FLG") = "I" Or UCase(oledbReader("DIR_FLG")) = "INCOMING" Or UCase(oledbReader("DIR_FLG")) = "INCOMMING" Then
                    CallType = "INCOMING"
                    A_Party = oledbReader("DLD_NO")
                    B_Party = oledbReader("DLG_NO")
                    IMEI = oledbReader("CLG_IMEI")
                ElseIf oledbReader("DIR_FLG") = "O" Or UCase(oledbReader("DIR_FLG")) = "OUTGOING" Then
                    CallType = "OUTGOING"
                    B_Party = oledbReader("DLD_NO")
                    A_Party = oledbReader("DLG_NO")
                    IMEI = oledbReader("CLD_IMIE")
                End If
                outFile.WriteLine(A_Party & " , " & B_Party & "," & DateToDouble(oledbReader("STRT_TM")) & "," & DateToDouble(oledbReader("END_TM")) & "," & CallType & "," & oledbReader("DRTN") & "," & IMEI & "," & IMSI & "," & oledbReader("Cell_id"))
                lbZongCount.Text = xlRowNumber
                lbZongCount.Refresh()
            End While
        Catch ex As Exception

        End Try
        oledbConn.Close()
        oledbReader.Close()
        SelectCommand.Dispose()
        prbZong.Value = prbZong.Value + 1
        prbZong.Refresh()
    End Sub
    Dim prbName As String
    Private Sub ProgressDisplay(ByVal prbName As String)
        If prbName = "prbMobilink" Then
        ElseIf prbName = "prbUfone" Then
        ElseIf prbName = "prbZong" Then
        ElseIf prbName = "prbTeleNor" Then

        End If
    End Sub
    Private Sub CreateCsvUfoneXLS(ByVal CellID_File_Path As String, ByVal SheetName As String)
        Dim oledbConn As OleDb.OleDbConnection
        Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & CellID_File_Path & ";extended properties='excel 12.0 xml; hdr=yes'"
        oledbConn = New OleDb.OleDbConnection(oledbconnstr)
        oledbConn.Open()
        Dim SelectCommand As New OleDb.OleDbCommand
        Dim oledbReader As OleDbDataReader
        SelectCommand.CommandText = "Select * from [" & SheetName & "$" & "]"
        SelectCommand.Connection = oledbConn
        Dim MobilinkFileName As String = Path.GetFileNameWithoutExtension(CellID_File_Path)
        Try
            oledbReader = SelectCommand.ExecuteReader()
            While oledbReader.Read()
                xlRowNumber = xlRowNumber + 1
                outFile.WriteLine(oledbReader("A_Number") & "," & oledbReader("B_Number") & "," & DateToDouble(oledbReader("Call_Start_Time")) & "," & DateToDouble(oledbReader("Call_End_Time")) & "," & oledbReader("CALL_INBOUND_OUTBOUND_DESC") & "," & oledbReader("Call_Duration") & "," & oledbReader("IMEI") & "," & oledbReader("IMSI") & "," & oledbReader("Cell_ID") & "," & oledbReader("LOCATION"))
                lbUfoneCount.Text = xlRowNumber
                lbUfoneCount.Refresh()
            End While
        Catch ex As Exception

        End Try
        oledbConn.Close()
        oledbReader.Close()
        SelectCommand.Dispose()
        prbUfone.Value = prbUfone.Value + 1
        prbUfone.Refresh()
    End Sub
    Private Sub CreateCsvUfone(ByVal CellID_File_Path As String)

        ' Using TextFieldParser to read tabular data
        Using sReader As New System.IO.StreamReader(CellID_File_Path)
            ' Set the delimiter (tab, comma, etc.) based on your file format

            Dim headers() As String = sReader.ReadLine().Split(vbTab)
            Dim colName As String = Nothing
            Dim ColIndex As Integer = 0
            Dim Indexs(9) As Integer
            For Each header In headers
                colName = headers(ColIndex)
                If colName = "A_Number" Then
                    Indexs(0) = ColIndex
                ElseIf colName = "B_Number" Then
                    Indexs(1) = ColIndex
                ElseIf colName = "Call_Start_Time" Then
                    Indexs(2) = ColIndex
                ElseIf colName = "Call_End_Time" Then
                    Indexs(3) = ColIndex
                ElseIf colName = "CALL_INBOUND_OUTBOUND_DESC" Then
                    Indexs(4) = ColIndex
                ElseIf colName = "Call_Duration" Then
                    Indexs(5) = ColIndex
                ElseIf colName = "IMEI" Then
                    Indexs(6) = ColIndex
                ElseIf colName = "IMSI" Then
                    Indexs(7) = ColIndex
                ElseIf colName = "Cell_ID" Then
                    Indexs(8) = ColIndex
                ElseIf colName = "LOCATION" Then
                    Indexs(9) = ColIndex
                End If
                ColIndex = ColIndex + 1
            Next
            ' Read the header row if your file has one
            'Dim headers As String() = sReader.ReadFields()
            Dim fields() As String
            ' Loop through the remaining rows
            While Not sReader.EndOfStream
                ' Read current row fields
                fields = sReader.ReadLine().Split(vbTab)

                ' Process or display the data as needed
                For i As Integer = 0 To headers.Length - 1

                Next
                xlRowNumber = xlRowNumber + 1
                'outFile.WriteLine(fields(3) & " , " & fields(4) & " , " & DateToDouble(Convert.ToDateTime(fields(5))) & " , " & DateToDouble(Convert.ToDateTime(fields(6))) & " , " & fields(9) & " , " & fields(7) & " , " & fields(1) & " , " & fields(2) & " , " & fields(8) & " , " & RemoveCommma(fields(10)))
                outFile.WriteLine(fields(Indexs(0)) & "," & fields(Indexs(1)) & "," & DateToDouble(Convert.ToDateTime(fields(Indexs(2)))) & "," & DateToDouble(Convert.ToDateTime(fields(Indexs(3)))) & "," & fields(Indexs(4)) & "," & fields(Indexs(5)) & "," & fields(Indexs(6)) & "," & fields(Indexs(7)) & "," & fields(Indexs(8)) & "," & RemoveCommma(fields(Indexs(9))))
                lbUfoneCount.Text = xlRowNumber
                lbUfoneCount.Refresh()

                ' You can also store the data in data structures or perform other operations
            End While
        End Using
        prbUfone.Value = prbUfone.Value + 1
        prbUfone.Refresh()
    End Sub

    Private Company As String = Nothing
    Private Sub CreateCsvTeleNor(ByVal CellID_File_Path As String)
        Using sReader As New System.IO.StreamReader(CellID_File_Path)
            ' Set the delimiter (tab, comma, etc.) based on your file format
            Dim CallType As String = Path.GetFileName(CellID_File_Path)
            If UCase(CallType).Contains("INCOMING") Then
                CallType = "INCOMING"
            ElseIf UCase(CallType).Contains("OUTGOING") Then
                CallType = "OUTGOING"
            End If
            Dim headers() As String = sReader.ReadLine().Split(","c)
            Dim colName As String = Nothing
            Dim ColIndex As Integer = 0
            Dim Indexs(8) As Integer
            If CallType = "INCOMING" Then
                For Each header In headers
                    colName = headers(ColIndex)
                    If colName = "MSISDN" Then
                        Indexs(0) = ColIndex
                    ElseIf colName = "CALL_ORIG_NUM" Then
                        Indexs(1) = ColIndex
                    ElseIf colName = "CALL_START_DT_TM" Then
                        Indexs(2) = ColIndex
                    ElseIf colName = "CALL_END_DT_TM" Then
                        Indexs(3) = ColIndex
                    ElseIf colName = "Call_Network_Volume" Then
                        Indexs(4) = ColIndex
                    ElseIf colName = "IMEI" Then
                        Indexs(5) = ColIndex
                    ElseIf colName = "IMSI" Then
                        Indexs(6) = ColIndex
                    ElseIf colName = "CELL_SITE_ID" Then
                        Indexs(7) = ColIndex
                    ElseIf colName = "LOCATION" Then
                        Indexs(8) = ColIndex
                    End If
                    ColIndex = ColIndex + 1
                Next
            ElseIf CallType = "OUTGOING" Then
                For Each header In headers
                    colName = headers(ColIndex)
                    If colName = "MSISDN" Then
                        Indexs(0) = ColIndex
                    ElseIf colName = "CALL_DIALED_NUM" Then
                        Indexs(1) = ColIndex
                    ElseIf colName = "CALL_START_DT_TM" Then
                        Indexs(2) = ColIndex
                    ElseIf colName = "CALL_END_DT_TM" Then
                        Indexs(3) = ColIndex
                    ElseIf colName = "Call_Network_Volume" Then
                        Indexs(4) = ColIndex
                    ElseIf colName = "IMEI" Then
                        Indexs(5) = ColIndex
                    ElseIf colName = "IMSI" Then
                        Indexs(6) = ColIndex
                    ElseIf colName = "CELL_SITE_ID" Then
                        Indexs(7) = ColIndex
                    ElseIf colName = "LOCATION" Then
                        Indexs(8) = ColIndex
                    End If
                    ColIndex = ColIndex + 1
                Next
            End If
            
            ' Read the header row if your file has one
            'Dim headers As String() = sReader.ReadFields()
            Dim fields() As String
            ' Loop through the remaining rows
            While Not sReader.EndOfStream
                ' Read current row fields
                fields = sReader.ReadLine().Split(","c)

                ' Process or display the data as needed
                For i As Integer = 0 To headers.Length - 1

                Next
                xlRowNumber = xlRowNumber + 1
                'outFile.WriteLine(fields(3) & " , " & fields(4) & " , " & DateToDouble(Convert.ToDateTime(fields(5))) & " , " & DateToDouble(Convert.ToDateTime(fields(6))) & " , " & fields(9) & " , " & fields(7) & " , " & fields(1) & " , " & fields(2) & " , " & fields(8) & " , " & RemoveCommma(fields(10)))
                outFile.WriteLine(fields(Indexs(0)) & " , " & fields(Indexs(1)) & " , " & DateToDouble(Convert.ToDateTime(fields(Indexs(2)))) & " , " & DateToDouble(Convert.ToDateTime(fields(Indexs(3)))) & " , " & CallType & " , " & fields(Indexs(4)) & " , " & fields(Indexs(5)) & " , " & fields(Indexs(6)) & " , " & fields(Indexs(7)) & " , " & RemoveCommma(fields(Indexs(8))))
                lbTeleNorCount.Text = xlRowNumber
                lbTeleNorCount.Refresh()

                ' You can also store the data in data structures or perform other operations
            End While
        End Using
        prbTeleNor.Value = prbTeleNor.Value + 1
        prbTeleNor.Refresh()
    End Sub
    Function Duration(ByVal StartTime As DateTime, ByVal EndTime As DateTime) As TimeSpan
        Return EndTime - StartTime
    End Function
    Function DateToDouble(ByVal CurrentDateTime As DateTime) As Double
        Return CurrentDateTime.ToOADate()
    End Function

    Function RemoveCommma(ByVal orginalString As String) As String
        Return orginalString.Replace(",", "")
    End Function
    Private Sub FormatXlFile(ByVal csvFile As String)
        Dim xlFilePath As String = Path.ChangeExtension(csvFile, ".xlsx")
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        xlWorkBook = xlApp.Workbooks.Open(xlFilePath)
        xlWorkSheet = xlWorkBook.ActiveSheet
        'Dim XlFile As String = Path.ChangeExtension(csvFile, ".xlsx")
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, 1)).EntireColumn.Cells.NumberFormat = "0"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 2), xlWorkSheet.Cells(1, 2)).EntireColumn.Cells.NumberFormat = "0"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 3), xlWorkSheet.Cells(1, 3)).EntireColumn.Cells.NumberFormat = "hh:MM:ss AM/PM"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 4), xlWorkSheet.Cells(1, 4)).EntireColumn.Cells.NumberFormat = "dd/mm/yyyy"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 5), xlWorkSheet.Cells(1, 5)).EntireColumn.Cells.NumberFormat = "@"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 6), xlWorkSheet.Cells(1, 6)).EntireColumn.Cells.NumberFormat = "0"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 7), xlWorkSheet.Cells(1, 7)).EntireColumn.Cells.NumberFormat = "0"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 8), xlWorkSheet.Cells(1, 9)).EntireColumn.Cells.NumberFormat = "0"
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 10), xlWorkSheet.Cells(1, 10)).EntireColumn.Cells.NumberFormat = "@"

        xlApp.DisplayAlerts = False
        Try
            xlWorkSheet.SaveAs(xlFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook)
            Try
                xlApp.Quit()
                xlWorkBook.Close()
                releaseObject(xlWorkBook)
            Catch ex As Exception

            End Try
            Try
                xlApp.Quit()
                releaseObject(xlApp)
                xlWorkBook.Close()
                releaseObject(xlWorkBook)

            Catch ex As Exception

            End Try
        Catch ex As Exception
            MsgBox("Error during formating excel file", MsgBoxStyle.OkOnly)
        End Try
        Call labelsText(Company, "Completed........")
    End Sub
    Private Function csvToExcel(ByVal csvFile As String)
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        xlWorkBook = xlApp.Workbooks.Open(csvFile)
        xlWorkSheet = xlWorkBook.ActiveSheet
        xlWorkSheet.Name = "bts"
        Dim XlFile As String = Path.ChangeExtension(csvFile, ".xlsx")
        xlApp.DisplayAlerts = False
        Try
            xlWorkSheet.SaveAs(XlFile, Excel.XlFileFormat.xlOpenXMLWorkbook)
            Try
                xlApp.Quit()
                xlWorkBook.Close()
                releaseObject(xlWorkBook)
            Catch ex As Exception

            End Try
            Try
                xlApp.Quit()
                releaseObject(xlApp)
                xlWorkBook.Close()
                releaseObject(xlWorkBook)

            Catch ex As Exception

            End Try
            If File.Exists(csvFile) Then
                ' Delete the file
                'File.Delete(csvFile)
                My.Computer.FileSystem.DeleteFile(csvFile, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently, FileIO.UICancelOption.DoNothing)
            End If
            'If Company = "Mobilink" Then
            '    lbMobilink.Text = "Finlazing ...... Step2"
            '    prbMbilink.Value = prbMbilink.Value + 1
            'ElseIf Company = "Ufone" Then
            '    lbUfone.Text = "Finlazing ...... Step2"
            '    prbUfone.Value = prbUfone.Value + 1
            'ElseIf Company = "TeleNor" Then
            '    lbTeleNor.Text = "Finlazing ...... Step2"
            '    prbTeleNor.Value = prbTeleNor.Value + 1
            'ElseIf Company = "Zong" Then
            '    prbZong.Value = prbZong.Value + 1
            '    lbZong.Text = "Finlazing ...... Step2"
            'End If
            labelsText(Company, "Finalizing ...... Step2")
            Call FormatXlFile(csvFile)
            'MsgBox("bts file has been created", MsgBoxStyle.OkOnly)

        Catch ex As Exception
            MsgBox("Error in converting csv to excel", MsgBoxStyle.OkOnly)
        End Try




    End Function
    Private Sub labelsText(ByVal CompanyName As String, ByVal LabelMsg As String)
        If Company = "Mobilink" Then
            prbMbilink.Value = prbMbilink.Value + 1
            prbMbilink.Refresh()
            lbMobilink.Text = LabelMsg
            lbMobilink.Refresh()
        ElseIf Company = "Ufone" Then
            lbUfone.Text = LabelMsg
            prbUfone.Value = prbUfone.Value + 1
            prbUfone.Refresh()
            lbUfone.Refresh()
        ElseIf Company = "TeleNor" Then
            lbTeleNor.Text = LabelMsg
            prbTeleNor.Value = prbTeleNor.Value + 1
            lbTeleNor.Refresh()
            prbTeleNor.Refresh()
        ElseIf Company = "Zong" Then
            prbZong.Value = prbZong.Value + 1
            lbZong.Text = LabelMsg
            lbZong.Refresh()
            prbZong.Refresh()
        End If
    End Sub
    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Call RefreshForm()
    End Sub
    Private CompletedFiles As String
    Private Sub RefreshForm()
        CompletedFiles = ""
        lbMobilink.Text = ""
        lbTeleNor.Text = ""
        lbUfone.Text = ""
        lbZong.Text = ""
        prbMbilink.Visible = False
        prbMbilink.Value = 0
        prbUfone.Visible = False
        prbUfone.Value = 0
        prbZong.Visible = False
        prbZong.Value = 0
        prbTeleNor.Visible = False
        prbTeleNor.Value = 0
        lbRecCount.Text = ""
        lbUfoneCount.Text = ""
        lbZongCount.Text = ""
        lbTeleNorCount.Text = ""
        txtFolderName.Text = ""
        isMobilink = False
        isUfone = False
        isZong = False
        isTeleNor = False
        lbMobilinkFileCount.Visible = False
        lbMobilinkFileCount.Text = ""
        lbUfoneFileCount.Visible = False
        lbUfoneFileCount.Text = ""
        lbZongFileCount.Text = ""
        lbZongFileCount.Visible = False
        lbTeleNorFileCount.Text = ""
        lbTeleNorFileCount.Visible = False
        btnMerge.Enabled = True

    End Sub

    Private Sub btnSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetting.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            'TextBox1.Text = FolderBrowserDialog1.SelectedPath
            'Try
            '    My.Computer.Registry.CurrentUser.DeleteSubKey("MegerPath", True)
            'Catch ex As Exception

            'End Try
            'Try
            '    My.Computer.Registry.CurrentUser.DeleteValue("MegerPath", True)
            'Catch ex As Exception

            'End Try
            ' My.Computer.Registry.CurrentUser.DeleteSubKey("UserName", True)

            'My.Computer.Registry.CurrentUser.DeleteValue("UserName", True)
            Dim regKey As Object = My.Computer.Registry.CurrentUser.GetValue("MegerPath")

            If regKey Is Nothing Then
                My.Computer.Registry.CurrentUser.CreateSubKey("MegerPath")
                My.Computer.Registry.CurrentUser.SetValue("MegerPath", FolderBrowserDialog1.SelectedPath)
            Else
                'My.Computer.Registry.CurrentUser.CreateSubKey("MegerPath")
                My.Computer.Registry.CurrentUser.SetValue("MegerPath", FolderBrowserDialog1.SelectedPath)
            End If

            
        End If
    End Sub

    Private Sub txtFolderName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFolderName.GotFocus
        If txtFolderName.Text = "Enter FIR Year Police Station" Then
            txtFolderName.Text = ""
            txtFolderName.ForeColor = Color.Black
        End If
    End Sub

    Private Sub txtFolderName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFolderName.LostFocus
        If txtFolderName.Text = "" Then
            txtFolderName.Text = "Enter FIR Year Police Station"
            txtFolderName.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub txtFolderName_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFolderName.MouseEnter
        If txtFolderName.Text = "Enter FIR Year Police Station" Then
            txtFolderName.Text = ""
            txtFolderName.ForeColor = Color.Black
        End If
    End Sub

    Private Sub txtFolderName_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFolderName.MouseLeave
        If txtFolderName.Text = "" Then
            txtFolderName.Text = "Enter FIR Year Police Station"
            txtFolderName.ForeColor = Color.Gray
        End If
    End Sub

   
    Private Sub txtFolderName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFolderName.TextChanged

    End Sub
End Class