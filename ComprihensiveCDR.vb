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
Imports Microsoft.VisualBasic.FileIO

Module ComprihensiveCDR
    Public Function mergeBs(ByVal excelFilePath As String, ByVal sName As String)
        Dim excelApp As New Excel.Application
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(excelFilePath)
        'xlWorkbook = xlWorkbooks.Open(excelFilePath)

        Dim SheetName As String = sName & "$"
        'xlDataSheet = xlWorkbook.Worksheets(sNamey
        Try
            Dim worksheet As Excel.Worksheet = workbook.Worksheets(sName)
            Dim range As Excel.Range = worksheet.Range("C:D")
            range.Merge()
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        Catch ex As Exception

        End Try
        
        workbook.Save()
        workbook.Close()
        excelApp.Quit()
    End Function
    Public Function IsFileCorrect(ByVal excelFilePath As String, ByVal SheetName As String) As String()
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlDataSheet As Excel.Worksheet = Nothing
        Dim misValue As Object = System.Reflection.Missing.Value
        frm_Spy_Tech.prgbar_CDR_report.Minimum = 0
        frm_Spy_Tech.prgbar_CDR_report.Maximum = 7
        frm_Spy_Tech.prgbar_CDR_report.Value = 0
        frm_Spy_Tech.prgbar_CDR_report.Visible = True
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim rng As Excel.Range
        xlWorkbook = xlWorkbooks.Open(excelFilePath)
        frm_Spy_Tech.prgbar_CDR_report.Value = 1
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        For Each xlSheet In xlWorkbook.Worksheets
            rng = xlSheet.UsedRange
            If rng.Count >= 99 Then
                xlSheet.Name = SheetName
                Exit For
            End If
        Next
        frm_Spy_Tech.prgbar_CDR_report.Value = 2
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim NoOfColumns As Integer = xlSheet.UsedRange.Columns.Count()
        Dim NoOfRows As Integer = xlSheet.UsedRange.Rows.Count()
        Dim colIndex As Integer = 0
        Dim rowsIndex As Integer = 0
        Dim CNIC As String = ""
        Dim aPartyNo As String = ""
        'Deleting the blank columns
        'For i As Integer = 1 To NoOfColumns
        '    colIndex = colIndex + 1
        '    'If xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).Value = Nothing Then
        '    '    xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Delete()
        '    '    colIndex = colIndex - 1
        '    'End If
        '    NoOfColumns = xlSheet.UsedRange.Columns.Count()
        'Next
        Dim isCorrect As Boolean = True
        For j As Integer = 1 To 4
            rowsIndex = rowsIndex + 1
            If CStr(xlSheet.Range(xlSheet.Cells(rowsIndex, 1), xlSheet.Cells(rowsIndex, 1)).Value) = "CNIC" Then
                CNIC = CStr(xlSheet.Range(xlSheet.Cells(rowsIndex, 2), xlSheet.Cells(rowsIndex, 2)).Value)
                isCorrect = False
            ElseIf CStr(xlSheet.Range(xlSheet.Cells(rowsIndex, 1), xlSheet.Cells(rowsIndex, 1)).Value) = "Mobile" Then
                aPartyNo = CStr(xlSheet.Range(xlSheet.Cells(rowsIndex, 2), xlSheet.Cells(rowsIndex, 2)).Value)
                isCorrect = False
            End If
        Next
        Dim aPartyInfo As String() = {aPartyNo, CNIC}
        frm_Spy_Tech.prgbar_CDR_report.Value = 3
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        NoOfColumns = xlSheet.UsedRange.Columns.Count()

        xlApp.DisplayAlerts = False
        'Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CompRpt_CallSammary")
        If isCorrect = False Then
            Dim rowsToDelete As Excel.Range = xlSheet.Range("1:4")
            rowsToDelete.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

            Try
                xlWorkbook.SaveAs(excelFilePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            Catch ex As Exception
                MsgBox("Error in deleting b1 and b2 columns " & Err.Description, MsgBoxStyle.Information)
            End Try
            Try
                xlWorkbook.Close()
            Catch ex As Exception

            End Try
            Try
                xlWorkbooks.Close()
                xlApp.Quit()
                releaseObject(xlSheet)
                'releaseObject(xlDataSheet)
                releaseObject(xlWorkbook)
                releaseObject(xlWorkbooks)
                releaseObject(xlApp)
            Catch ex As Exception

            End Try
        ElseIf isCorrect = True Then
            'If NoOfColumns <= 3 Then
            '    MsgBox("If the company is Zong then Delete the all top rows of columns heading row", MsgBoxStyle.Information)
            '    isCorrect = False
            'End If
            Try
                xlWorkbook.Close(SaveChanges:=False)
                xlWorkbooks.Close()
            Catch ex As Exception
            End Try
            Try
                xlApp.Quit()
                releaseObject(xlSheet)
                'releaseObject(xlDataSheet)
                releaseObject(xlWorkbook)
                releaseObject(xlWorkbooks)
                releaseObject(xlApp)
            Catch ex As Exception

            End Try
        End If

        frm_Spy_Tech.prgbar_CDR_report.Value = 4
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Return aPartyInfo
    End Function

    Public Function ColumnsName()
        Dim OthersCon As SqlConnection
        OthersCon = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If OthersCon.State = ConnectionState.Closed Then
            OthersCon.Open()
        End If
        Dim cmd As New SqlCommand("IF OBJECT_ID('tblFormatCDRcols', 'U') IS NOT NULL SELECT 1 ELSE SELECT 0", OthersCon)
        'Dim cmd As New SqlCommand("IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'tableName') SELECT 1 ELSE SELECT 0", OthersCon)
        Dim tableExists As Integer = CInt(cmd.ExecuteScalar())
        If tableExists = 0 Then
            Dim dtTable As New Data.DataTable
            dtTable.Columns.Add("OldName", GetType(String))
            dtTable.Columns.Add("NewName", GetType(String))
            Dim row As DataRow = dtTable.NewRow()
            dtTable.Rows.Add("A Number", "a")
            dtTable.Rows.Add("A-Party", "a")
            dtTable.Rows.Add("B Number", "b")
            dtTable.Rows.Add("BNUMBER", "b")
            dtTable.Rows.Add("B-Party", "b")
            dtTable.Rows.Add("Call Type", "Call Type")
            dtTable.Rows.Add("CALL_DIALED_NUM", "b2")
            dtTable.Rows.Add("Call_Network_Volume", "Duration")
            dtTable.Rows.Add("call_org_num", "b1")
            dtTable.Rows.Add("CALL_START_DT_TM", "Date & Time")
            dtTable.Rows.Add("CALL_TYPE", "Call Type")
            dtTable.Rows.Add("Cell Id", "Cell ID")
            dtTable.Rows.Add("CELL_ID", "Cell ID")
            dtTable.Rows.Add("Cell_SITE_ID", "Cell ID")
            dtTable.Rows.Add("Date & Time", "Date & Time")
            dtTable.Rows.Add("Direction", "Call Type")
            dtTable.Rows.Add("Duration", "Duration")
            dtTable.Rows.Add("INBOUND_OUTBOUND_IND", "Call Type")
            dtTable.Rows.Add("Location", "Site")
            dtTable.Rows.Add("MINS", "DurationMIN")
            dtTable.Rows.Add("MSISDN", "a")
            dtTable.Rows.Add("MSISDN_ID", "a")
            dtTable.Rows.Add("SECS", "DurationSECS")
            dtTable.Rows.Add("SITE_ADDRESS", "Site")
            dtTable.Rows.Add("Start Time", "Date & Time")
            dtTable.Rows.Add("STRT_TM", "Date & Time")

            Dim cmd1 As New SqlCommand("CREATE TABLE tblFormatCDRcols (OldName nvarchar(30), NewName nvarchar(30))", OthersCon)
            cmd1.ExecuteNonQuery()
            OthersCon.Close()
            OthersCon.Open()
            Dim bulkCopy As New SqlBulkCopy(OthersCon)
            bulkCopy.DestinationTableName = "tblFormatCDRcols"
            bulkCopy.WriteToServer(dtTable)
            cmd.Dispose()
            cmd1.Dispose()
            OthersCon.Close()
        Else

        End If
    End Function
    Public Function ChangeColsName(ByVal excelFilePath As String, ByVal SheetName As String)
        Call ColumnsName()
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlDataSheet As Excel.Worksheet = Nothing
        Dim misValue As Object = System.Reflection.Missing.Value
        frm_Spy_Tech.prgbar_CDR_report.Minimum = 0
        frm_Spy_Tech.prgbar_CDR_report.Maximum = 7
        frm_Spy_Tech.prgbar_CDR_report.Value = 0
        frm_Spy_Tech.prgbar_CDR_report.Visible = True
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        frm_Spy_Tech.prgbar_CDR_report.Value = 1
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim rng As Excel.Range
        xlWorkbook = xlWorkbooks.Open(excelFilePath)
        For Each xlSheet In xlWorkbook.Worksheets
            rng = xlSheet.UsedRange
            If rng.Count >= 99 Then
                xlSheet.Name = SheetName
                Exit For
            End If
        Next
        frm_Spy_Tech.prgbar_CDR_report.Value = 2
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim NoOfColumns As Integer = xlSheet.UsedRange.Columns.Count()

        Dim colIndex As Integer = 0
        Dim OthersCon As SqlConnection
        Dim cmdColName As SqlCommand
        Dim ColNameReader As SqlDataReader
        Dim ColNameDA As New SqlDataAdapter
        'Deleting the blank columns
        OthersCon = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If OthersCon.State = ConnectionState.Closed Then
            OthersCon.Open()
        End If
        frm_Spy_Tech.prgbar_CDR_report.Value = 3
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim queryString As String
        For i As Integer = 1 To NoOfColumns
            colIndex = colIndex + 1
            'Try
            '    xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Cells.NumberFormat = "0"
            'Catch ex As Exception

            'End Try

            If xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).Value = Nothing Then
                xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Delete()
                colIndex = colIndex - 1
            End If
            NoOfColumns = xlSheet.UsedRange.Columns.Count()
        Next
        frm_Spy_Tech.prgbar_CDR_report.Value = 4
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim OldcolName As String

        NoOfColumns = xlSheet.UsedRange.Columns.Count()
        colIndex = 0
        Dim IsBs As Boolean = False
        For i As Integer = 1 To NoOfColumns
            colIndex = colIndex + 1
            If xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).Value <> Nothing Then
                OldcolName = xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).Value
                queryString = "Select NewName from tblFormatCDRcols where OldName = '" & Trim(OldcolName.Replace(vbCrLf, "")) & "'"
                cmdColName = New SqlCommand(queryString, OthersCon)
                Try
                    ColNameReader = cmdColName.ExecuteReader
                    While (ColNameReader.Read)
                        xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).Value = ColNameReader(0).ToString
                        If ColNameReader(0).ToString = "a" Then
                            Try
                                xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Cells.NumberFormat = "0"
                            Catch ex As Exception

                            End Try
                        ElseIf ColNameReader(0).ToString = "b" Then
                            Try
                                xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Cells.NumberFormat = "0"
                            Catch ex As Exception

                            End Try
                        ElseIf ColNameReader(0).ToString = "b1" Then
                            Try
                                xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Cells.NumberFormat = "0"
                            Catch ex As Exception

                            End Try
                            IsBs = True
                            xlSheet.Columns(i).Insert()
                            xlSheet.Range(xlSheet.Cells(1, i), xlSheet.Cells(1, i)).EntireColumn.Cells.NumberFormat = "@"
                            xlSheet.Cells(1, i).value2 = "b"

                            ' xlSheet.Range(xlSheet.Cells(1, i), xlSheet.Cells(1, i)).EntireColumn.ColumnWidth = "255"
                        ElseIf ColNameReader(0).ToString = "b2" Then
                            Try
                                xlSheet.Range(xlSheet.Cells(1, colIndex), xlSheet.Cells(1, colIndex)).EntireColumn.Cells.NumberFormat = "0"
                            Catch ex As Exception

                            End Try
                        End If

                        Exit While
                    End While
                    ColNameReader.Close()
                Catch ex As Exception

                End Try
            End If
            NoOfColumns = xlSheet.UsedRange.Columns.Count()
        Next

        Try
            OthersCon.Close()
            ColNameReader.Close()
            frm_Spy_Tech.prgbar_CDR_report.Value = 5
            frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Catch ex As Exception

        End Try

        xlApp.DisplayAlerts = False
        'Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CompRpt_CallSammary")
        Try
            xlWorkbook.SaveAs(excelFilePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        Catch ex As Exception
            MsgBox("Error in saving Comprehaensive CDR Report " & Err.Description, MsgBoxStyle.Information)
        End Try
        xlWorkbook.Close()
        xlWorkbooks.Close()
        xlApp.Quit()
        releaseObject(xlSheet)
        'releaseObject(xlDataSheet)
        releaseObject(xlWorkbook)
        releaseObject(xlWorkbooks)
        releaseObject(xlApp)
        frm_Spy_Tech.prgbar_CDR_report.Value = 6
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        If IsBs = True Then

            'Call ColumnsName()
            Dim oledbConn As OleDb.OleDbConnection
            Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & excelFilePath & ";extended properties='excel 12.0 xml; hdr=yes'"
            oledbConn = New OleDb.OleDbConnection(oledbconnstr)
            oledbConn.Open()
            'Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & XLfilePath & ";extended properties='excel 12.0 xml; hdr=yes'"
            Dim InsertCommand As New OleDb.OleDbCommand
            Dim DelCommand As New OleDb.OleDbCommand
            'Dim DelReader As New OleDbDataReader
            'InsertCommand = New OleDb.OleDbCommand(queryString, oledbConn)
            Try
                InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b] = [b2]  Where [Call Type] = 'OUTGOING' and LEN(b2)<=30"
                InsertCommand.Connection = oledbConn
                InsertCommand.ExecuteNonQuery()
                InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b] = [b2]  Where [Call Type] = 'DATA' and LEN(b2)<=30"
                InsertCommand.Connection = oledbConn
                InsertCommand.ExecuteNonQuery()
                InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b] = [b1]  Where [Call Type] = 'INCOMING' and LEN(b1)<=30"
                InsertCommand.Connection = oledbConn
                InsertCommand.ExecuteNonQuery()
                InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b1] = NULL  Where [Call Type] = 'OUTGOING'" ' and LEN(b2)<=30"
                InsertCommand.Connection = oledbConn
                InsertCommand.ExecuteNonQuery()
                InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b1] = NULL  Where [Call Type] = 'DATA'" ' and LEN(b2)<=30"
                InsertCommand.Connection = oledbConn
                InsertCommand.ExecuteNonQuery()
                InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b2] = NULL  Where [Call Type] = 'INCOMING'" ' and LEN(b1)<=30"
                InsertCommand.Connection = oledbConn
                InsertCommand.ExecuteNonQuery()
                'InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b] = [b1] &''& [b2]"
                'InsertCommand.Connection = oledbConn
                'InsertCommand.ExecuteNonQuery()
                'DelCommand.CommandText = "Select Distinct a From [" & SheetName & "$" & "]"
                'DelCommand.Connection = oledbConn
                'Dim DelReader As OleDb.OleDbDataReader = DelCommand.ExecuteReader()
                'While DelReader.Read()
                '    InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b1] = ''  Where [b1] = '" & DelReader(0) & "'"
                '    InsertCommand.Connection = oledbConn
                '    InsertCommand.ExecuteNonQuery()

                'End While
                'DelCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b1] = ""  Where [b1] <> [b1] "
                'InsertCommand.Connection = oledbConn
                'InsertCommand.ExecuteNonQuery()
                'InsertCommand.CommandText = "Update [" & SheetName & "$" & "] SET [b] = [b2]  Where [a] <> [b2] and LEN(b2)<=20"
                'InsertCommand.Connection = oledbConn
                'InsertCommand.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            InsertCommand.Dispose()
            oledbConn.Close()
        End If
        'Call mergeBs(excelFilePath, SheetName)
        Call DelBs(excelFilePath, SheetName)
        frm_Spy_Tech.prgbar_CDR_report.Value = 7
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
    End Function
    Function DelBs(ByVal excelFilePath As String, ByVal sName As String)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlDataSheet As Excel.Worksheet = Nothing
        xlApp = New Excel.Application
        Dim range As Excel.Range
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Dim SheetName As String = Nothing

        xlWorkbook = xlWorkbooks.Open(excelFilePath)

        SheetName = sName & "$"
        xlDataSheet = xlWorkbook.Worksheets(sName)

        Dim NoOfColumns As Integer = xlDataSheet.UsedRange.Columns.Count()
        Dim colIndex As Integer = 0

        'Deleting the blank columns
        Dim NoDelCols As Integer = 0
        For i As Integer = 1 To NoOfColumns
            colIndex = colIndex + 1
            If (xlDataSheet.Range(xlDataSheet.Cells(1, colIndex), xlDataSheet.Cells(1, colIndex)).Value = "b1") Or (xlDataSheet.Range(xlDataSheet.Cells(1, colIndex), xlDataSheet.Cells(1, colIndex)).Value = "b2") Then
                xlDataSheet.Range(xlDataSheet.Cells(1, colIndex), xlDataSheet.Cells(1, colIndex)).EntireColumn.Delete()
                colIndex = colIndex - 1
                NoDelCols = NoDelCols + 1
            End If
            If NoDelCols = 2 Then
                Exit For
            End If
            NoOfColumns = xlDataSheet.UsedRange.Columns.Count()
        Next
        xlApp.DisplayAlerts = False
        'Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CompRpt_CallSammary")
        Try
            xlWorkbook.SaveAs(excelFilePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        Catch ex As Exception
            MsgBox("Error in deleting b1 and b2 columns " & Err.Description, MsgBoxStyle.Information)
        End Try
        Try
            xlWorkbook.Close()
        Catch ex As Exception

        End Try
        Try
            xlWorkbooks.Close()
            xlApp.Quit()
            releaseObject(xlSheet)
            'releaseObject(xlDataSheet)
            releaseObject(xlWorkbook)
            releaseObject(xlWorkbooks)
            releaseObject(xlApp)
        Catch ex As Exception

        End Try

    End Function
    Public Function InsertColsCNIC(ByVal filesPath As String) As String
        Dim SaveFilePath As String = filesPath
        Dim Xl_file As String
        'Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CompRpt_CallSammary")
        'Call AutoFitExcelFile(Xl_file)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlDataSheet As Excel.Worksheet = Nothing
        frm_Spy_Tech.prgbar_CDR_report.Minimum = 0
        frm_Spy_Tech.prgbar_CDR_report.Maximum = 6
        frm_Spy_Tech.prgbar_CDR_report.Value = 0
        frm_Spy_Tech.prgbar_CDR_report.Visible = True
        xlApp = New Excel.Application
        Dim range As Excel.Range
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Dim SheetName As String = Nothing
        Dim TableName As String = "tblCallSummary"

        Try
            xlWorkbook = xlWorkbooks.Open(filesPath)
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "bts" Then
                    SheetName = "bts$"
                    xlDataSheet = xlWorkbook.Worksheets("bts")
                    IsSheetRenamed = True
                    Exit For
                ElseIf xlSheet.Name = "cdr" Then
                    SheetName = "cdr$"
                    xlDataSheet = xlWorkbook.Worksheets("cdr")
                    IsSheetRenamed = True
                    Exit For
                End If

            Next xlSheet
            frm_Spy_Tech.prgbar_CDR_report.Value = 1
            frm_Spy_Tech.prgbar_CDR_report.Refresh()
            If IsSheetRenamed = True Then
                Dim NoOfColumns As Integer = xlDataSheet.UsedRange.Columns.Count()
                Dim colIndex As Integer = 0

                'Deleting the blank columns

                For i As Integer = 1 To NoOfColumns
                    colIndex = colIndex + 1
                    If xlDataSheet.Range(xlDataSheet.Cells(1, colIndex), xlDataSheet.Cells(1, colIndex)).Value = Nothing Then
                        xlDataSheet.Range(xlDataSheet.Cells(1, colIndex), xlDataSheet.Cells(1, colIndex)).EntireColumn.Delete()
                        colIndex = colIndex - 1
                    End If
                    NoOfColumns = xlDataSheet.UsedRange.Columns.Count()
                Next
                NoOfColumns = xlDataSheet.UsedRange.Columns.Count()
                Dim isCNIC_a As Boolean = False
                Dim isCNIC_b As Boolean = False
                Dim isboth_a_b As Boolean = False

                ' Inserting columns for a and b parties
                frm_Spy_Tech.prgbar_CDR_report.Value = 2
                frm_Spy_Tech.prgbar_CDR_report.Refresh()
                For i As Integer = 1 To NoOfColumns
                    If xlDataSheet.Range(xlDataSheet.Cells(1, i), xlDataSheet.Cells(1, i)).Value = "a" Then
                        If xlDataSheet.Range(xlDataSheet.Cells(1, i + 1), xlDataSheet.Cells(1, i + 1)).Value = Nothing Then
                            xlDataSheet.Range(xlDataSheet.Cells(1, i + 1), xlDataSheet.Cells(1, i + 1)).EntireColumn.Delete()
                        End If
                        xlDataSheet.Columns(i + 1).Insert()
                        xlDataSheet.Cells(1, i + 1).value2 = "CNIC_a"
                        xlDataSheet.Range(xlDataSheet.Cells(1, i + 1), xlDataSheet.Cells(1, i + 1)).EntireColumn.NumberFormat = "@"
                        isCNIC_a = True
                    End If
                    If xlDataSheet.Range(xlDataSheet.Cells(1, i), xlDataSheet.Cells(1, i)).Value = "b" Then
                        If xlDataSheet.Range(xlDataSheet.Cells(1, i + 1), xlDataSheet.Cells(1, i + 1)).Value = Nothing Then
                            xlDataSheet.Range(xlDataSheet.Cells(1, i + 1), xlDataSheet.Cells(1, i + 1)).EntireColumn.Delete()
                        End If
                        xlDataSheet.Columns(i + 1).Insert()
                        xlDataSheet.Cells(1, i + 1).value2 = "CNIC_b"
                        xlDataSheet.Range(xlDataSheet.Cells(1, i + 1), xlDataSheet.Cells(1, i + 1)).EntireColumn.NumberFormat = "@"
                        isCNIC_b = True
                    End If
                    If isCNIC_a = True And isCNIC_b = True Then
                        'If isboth_a_b = False Then
                        '    isboth_a_b = True
                        'ElseIf isboth_a_b = True Then
                        Exit For
                        'End If

                    End If
                Next
                frm_Spy_Tech.prgbar_CDR_report.Value = 3
                frm_Spy_Tech.prgbar_CDR_report.Refresh()
                xlDataSheet.UsedRange.EntireColumn.AutoFit()
                xlApp.DisplayAlerts = False
                Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CompRpt_CallSammary")
                Try
                    xlWorkbook.SaveAs(Xl_file, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
                Catch ex As Exception
                    MsgBox("Error in saving Comprehaensive CDR Report " & Err.Description, MsgBoxStyle.Information)
                End Try

                'MsgBox("File has been created: " & Xl_file, MsgBoxStyle.Information)
                xlWorkbook.Close()
                xlWorkbooks.Close()
                xlApp.Quit()
                releaseObject(xlSheet)
                releaseObject(xlDataSheet)
                releaseObject(xlWorkbook)
                releaseObject(xlWorkbooks)
                releaseObject(xlApp)

            End If
            fill_Temp(Xl_file, SheetName)
            frm_Spy_Tech.prgbar_CDR_report.Value = 4
            frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Catch ex As Exception

            MsgBox("Error in creating Comprehensive CDR Report " & Err.Description, MsgBoxStyle.Information)
            Try
                xlWorkbook.Close()
                xlWorkbooks.Close()
                xlApp.Quit()
            Catch ex2 As Exception

            End Try

            Try
                releaseObject(xlSheet)
                releaseObject(xlDataSheet)
                releaseObject(xlWorkbook)
                releaseObject(xlWorkbooks)
                releaseObject(xlApp)
            Catch ex3 As Exception

            End Try

        End Try
        Call AutoFitExcelFile(Xl_file)
        frm_Spy_Tech.prgbar_CDR_report.Value = 5
        frm_Spy_Tech.prgbar_CDR_report.Refresh()

        Return Xl_file
    End Function
    Public subsInfo As String()
    Sub fill_Temp(ByVal XLfilePath As String, ByVal SheetName As String)
        frm_Spy_Tech.prgbar_CDR_report.Minimum = 0
        frm_Spy_Tech.prgbar_CDR_report.Maximum = 5
        frm_Spy_Tech.prgbar_CDR_report.Value = 0
        frm_Spy_Tech.prgbar_CDR_report.Visible = True
        Dim sqlconn As SqlConnection
        Dim CreateTbCommand As SqlCommand
        Dim oledbConn As OleDb.OleDbConnection
        Dim InsertCommand As OleDbCommand
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim ColumnName As String = "b"
        'Dim SheetName As String = "cdr$"
        Dim totalRows As Integer = 0
        Dim TempTable As String = "tblXLNums"
        Dim db2021Connection As SqlConnection
        Dim OthersCon As SqlConnection
        'Create DataTable for results
        OthersCon = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If OthersCon.State = ConnectionState.Closed Then
            OthersCon.Open()
        End If
        frm_Spy_Tech.prgbar_CDR_report.Value = 1
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim TargetConnection As SqlConnection
        TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If OthersCon.State = ConnectionState.Closed Then
            TargetConnection.Open()
        End If
        db2021Connection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;")
        If OthersCon.State = ConnectionState.Closed Then
            db2021Connection.Open()
        End If
        Try
            Dim TempQrStr As String = "Drop Table IF EXISTS [" & TempTable & "]"
            CreateTbCommand = New SqlCommand(TempQrStr, OthersCon)
            CreateTbCommand.ExecuteNonQuery()
        Catch ex03 As Exception
        End Try
        frm_Spy_Tech.prgbar_CDR_report.Value = 2
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim CreateTablesQuery As String = "CREATE TABLE [" & TempTable & "] (PhoneNumber nvarchar(255) null, UpdatedNumbers nvarchar(255) null)"

        '*** Distinct phone Numbers to Sql Server *****

        OthersCommand = New SqlCommand(CreateTablesQuery, OthersCon)
        Try
            OthersCommand.ExecuteNonQuery()
            'targetconnection.close()
        Catch ex02 As Exception
            MsgBox("creating table", MsgBoxStyle.OkOnly)
            'targetconnection.close()
        End Try
        'Dim exceltosql_query As String = "select b , [call type] from [" & SheetName & "]"
        '92'+trim('0' from phonenumber) as party from tblxlnums s where s.phonenumber like '[0]%' and s.phonenumber not like '[0092]'
        'select party from 
        Dim querystring As String = "select distinct [" & ColumnName1 & "] as party from [" & SheetName & "] where isnumeric(" & ColumnName1 & ") and len(" & ColumnName1 & ") >= 10 " + _
                        "union all select distinct [" & ColumnName2 & "] as party from [" & SheetName & "] where isnumeric(" & ColumnName2 & ") and len(" & ColumnName2 & ") >= 10"
        Dim oledbconnstr As String = "provider=microsoft.ace.oledb.12.0;data source=" & XLfilePath & ";extended properties='excel 12.0 xml; hdr=yes'"
        oledbConn = New OleDb.OleDbConnection(oledbconnstr)
        oledbConn.Open()
        InsertCommand = New OleDb.OleDbCommand(querystring, oledbConn)
        Dim insertreader As OleDb.OleDbDataReader
        Try
            insertreader = InsertCommand.ExecuteReader()
        Catch ex02 As Exception
            MsgBox("insertion error table", MsgBoxStyle.OkOnly)
        End Try
        frm_Spy_Tech.prgbar_CDR_report.Value = 3
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        '*** Excel Sheet to Sql Server ****
        'Dim querystringAllRecs As String = "select * from [" & SheetName & "]"
        'Dim oledbconnstring As String = "provider=microsoft.ace.oledb.12.0;data source=" & XLfilePath & ";extended properties='excel 12.0 xml;'" 'hdr=0;imex=0'"
        'oledbConn = New OleDb.OleDbConnection(oledbconnstring)
        'oledbConn.Open()
        'InsertCommand = New OleDb.OleDbCommand(querystringAllRecs, oledbConn)
        'Dim insertreader As OleDb.OleDbDataReader
        'CreateTablesQuery = "CREATE TABLE [" & TempTable & "]"
        'Try
        '    Dim dset As DataSet
        '    'insertreader = insertcommand.executereader()
        '    Dim dataadapter As OleDbDataAdapter = New OleDbDataAdapter(InsertCommand)
        '    'dataadapter.Fill(dset)
        '    Dim dt As Data.DataTable = New Data.DataTable
        '    dataadapter.Fill(dt)
        '    Dim firstcol As Boolean = True
        '    For Each datacols As DataColumn In dt.Columns
        '        If firstcol = True Then
        '            CreateTablesQuery = CreateTablesQuery + " ([" & datacols.ColumnName & "] nvarchar(255) null"
        '            firstcol = False
        '        ElseIf firstcol = False Then
        '            CreateTablesQuery = CreateTablesQuery + " ,[" & datacols.ColumnName & "] nvarchar(255) null"
        '        End If

        '    Next
        '    CreateTablesQuery = CreateTablesQuery + " )"
        '    OthersCommand = New SqlCommand(CreateTablesQuery, OthersCon)
        '    Try
        '        OthersCommand.ExecuteNonQuery()
        '        'TargetConnection.Close()
        '    Catch ex02 As Exception
        '        MsgBox("creating table", MsgBoxStyle.OkOnly)
        '        'TargetConnection.Close()
        '    End Try
        'Catch ex02 As Exception
        '    MsgBox("insertion error table", MsgBoxStyle.OkOnly)
        'End Try

        'Try
        '    insertreader = InsertCommand.ExecuteReader()
        'Catch ex02 As Exception
        '    MsgBox("insertion error table", MsgBoxStyle.OkOnly)
        'End Try
        '*** Excel Sheet to Sql Server ended ****

        Dim bcCopy As New SqlBulkCopy(OthersCon)
        'TargetConnection.Open()
        bcCopy.BatchSize = 100000
        bcCopy.BulkCopyTimeout = 0
        bcCopy.DestinationTableName = "[" & TempTable & "]"
        bcCopy.WriteToServer(insertreader)
        insertreader.Close()
        frm_Spy_Tech.prgbar_CDR_report.Value = 4
        frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Dim updateQry(2) As String
        'updateQry(0) = "Update tblXLNums set PhoneNumber = REPLACE(LTRIM(REPLACE(PhoneNumber, '0', ' ')), ' ', '0') where PhoneNumber like '[09]%' OR PhoneNumber like '[009]%'"
        updateQry(0) = "Update tblXLNums set UpdatedNumbers =  REPLACE(LTRIM(REPLACE(PhoneNumber, '0', ' ')), ' ', '0') where (PhoneNumber like '09%') OR (PhoneNumber like '009%') OR (PhoneNumber like '92%')"
        updateQry(1) = "Update tblXLNums set UpdatedNumbers = '92' + REPLACE(LTRIM(REPLACE(PhoneNumber, '0', ' ')), ' ', '0') where (PhoneNumber like '0%')"
        updateQry(2) = "Update tblXLNums set UpdatedNumbers = '92' + PhoneNumber where (LEN(PhoneNumber) = 10)"
        Dim ValueEffected As Integer = 0
        Try
            For Each Query As String In updateQry
                'updateQry = updateQry
                CreateTbCommand = New SqlCommand(Query, OthersCon)
                ValueEffected = CreateTbCommand.ExecuteNonQuery()
                ' MsgBox("Effected Rows are " & ValueEffected)
            Next
        Catch ex03 As Exception
        End Try
        Dim ViewCNICqueryStr As String = "Select * From ViewCNIC"
        Dim ViewReader As SqlDataReader
        Dim ViewCommand As SqlCommand
        ViewCommand = New SqlCommand(ViewCNICqueryStr, OthersCon)
        Dim updateXLquery As String
        Dim PhoneNO, CNIC As String

        Try
            ViewReader = ViewCommand.ExecuteReader()
            'InsertCommand.CommandText = "Update [" & SheetName & "] set [b] = '92' + REPLACE(LTRIM(REPLACE(b, '0', ' ')), ' ', '0') where (b like '[0]%') AND (b not like '009%')"
            'InsertCommand.Connection = oledbConn
            'InsertCommand.ExecuteNonQuery()
            While ViewReader.Read()
                PhoneNO = ViewReader(0).ToString()
                'If ViewReader.IsDBNull(1) = True And ViewReader.IsDBNull(2) = True Then
                '    'CNIC = ViewReader(1).ToString()
                '    Continue While
                'ElseIf ViewReader.IsDBNull(1) = False Then
                '    CNIC = ViewReader(1).ToString()
                'ElseIf ViewReader.IsDBNull(2) = False Then
                CNIC = ViewReader(1).ToString()

                'End If
                'InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_a] =? Where Trim([a]) =?"

                'With InsertCommand.Parameters
                '    .AddWithValue("?", CNIC)
                '    .AddWithValue("?", PhoneNO)
                'End With
                'InsertCommand = New OleDb.OleDbCommand(updateXLquery, oledbConn)
                Try

                Catch ex As Exception

                End Try
                Try
                    If subsInfo(0) <> "" And subsInfo(1) <> "" Then
                        InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_a] = '" & subsInfo(1) & "' Where Trim([a]) = '" & subsInfo(0) & "'"
                        InsertCommand.Connection = oledbConn
                        InsertCommand.ExecuteNonQuery()
                    Else
                        InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_a] = '" & CNIC & "' Where Trim([a]) = '" & PhoneNO & "'"
                        InsertCommand.Connection = oledbConn
                        InsertCommand.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                End Try
                Try
                    If subsInfo(0) <> "" And subsInfo(1) <> "" Then
                        InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_a] = '" & subsInfo(1) & "' Where Trim([a]) = " & subsInfo(0) & ""
                        InsertCommand.Connection = oledbConn
                        InsertCommand.ExecuteNonQuery()
                    Else
                        InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_a] = '" & CNIC & "' Where Trim([a]) = " & PhoneNO & ""
                        InsertCommand.Connection = oledbConn
                        InsertCommand.ExecuteNonQuery()
                    End If
                Catch ex As Exception

                End Try
                'Exit While
                'InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_b] =? Where Trim([b]) =?"

                'With InsertCommand.Parameters
                '    .AddWithValue("?", CNIC)
                '    .AddWithValue("?", PhoneNO)
                'End With
                'InsertCommand = New OleDb.OleDbCommand(updateXLquery, oledbConn)
                Try
                    InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_b] = '" & CNIC & "' Where Trim([b]) = '" & PhoneNO & "'"
                    InsertCommand.Connection = oledbConn
                    InsertCommand.ExecuteNonQuery()
                Catch ex As Exception

                End Try
                Try
                    InsertCommand.CommandText = "Update [" & SheetName & "] SET [CNIC_b] = '" & CNIC & "' Where Trim([b]) = " & PhoneNO & ""
                    InsertCommand.Connection = oledbConn
                    InsertCommand.ExecuteNonQuery()
                Catch ex As Exception

                End Try
                'Exit While
            End While
            frm_Spy_Tech.prgbar_CDR_report.Value = 5
            frm_Spy_Tech.prgbar_CDR_report.Refresh()
        Catch ex As Exception
            ViewReader.Close()
            oledbConn.Close()
        End Try
        ViewReader.Close()
        oledbConn.Close()

        'Try

        '    CreateTbCommand = New SqlCommand(updateQry2, OthersCon)
        '    EffectRows = CreateTbCommand.ExecuteNonQuery()
        '    MsgBox("Effected Rows are " + EffectRows)
        'Catch ex03 As Exception
        'End Try
        'MsgBox("Data transfered successfully")

    End Sub
    Sub AutoFitExcelFile(ByVal FileXLAddress As String)

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing

        xlApp = New Excel.Application
        Try
            xlWorkbook = xlApp.Workbooks.Open(FileXLAddress)
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "cdr" Then
                    xlSheet = xlWorkbook.Worksheets("cdr")
                    xlSheet.Name = "Comprehensive_cdrRPT"
                    Exit For
                ElseIf xlSheet.Name = "bts" Then
                    xlSheet = xlWorkbook.Worksheets("bts")
                    xlSheet.Name = "Comprehensive_btsRPT"

                    Exit For
                End If
            Next xlSheet
            'xlSheet = xlWorkbook.Worksheets(SheetName)
            xlSheet.UsedRange.EntireColumn.AutoFit()
            xlWorkbook.Save()    ', misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        Catch ex As Exception

        End Try
        Try
            xlWorkbook.Close()
        Catch ex As Exception

        End Try

        Try
            xlWorkbooks.Close()
        Catch ex As Exception

        End Try

        xlApp.Quit()
        releaseObject(xlSheet)

        releaseObject(xlWorkbook)
        releaseObject(xlWorkbooks)
        releaseObject(xlApp)
    End Sub
    'Function Excel_To_SQL(ByVal CdrFile As String, ByVal onlyXLFilePath As String)
    'Dim ColumnName1 As String = "a"
    'Dim ColumnName2 As String = "b"
    'Dim SheetName As String = "cdr$"
    ''Dim TotalNumberOfCDRs As Integer = AllCommonFileNames.Length()
    ''Dim OnlyFileNames(TotalNumberOfCDRs) As String
    ' ''frm_Spy_Tech.prgbar_Common_Links.Minimum = 0
    '' ''frm_Spy_Tech.prgbar_Common_Links.Maximum = AllCommonFileNames.Length
    ' ''frm_Spy_Tech.prgbar_Common_Links.Value = 0
    ' ''frm_Spy_Tech.prgbar_Common_Links.Visible = True
    ' ''frm_Spy_Tech.prgbar_Common_Links.Refresh()
    ''Dim TransferedFiles As Integer = 0
    ''frm_Spy_Tech.lbCommonNos.Text = "Excel to SQL " & TransferedFiles.ToString & " of " & TotalNumberOfCDRs.ToString
    ''frm_Spy_Tech.lbCommonNos.Visible = True
    ''frm_Spy_Tech.lbCommonNos.Refresh()
    'Dim db2021Connection As SqlConnection
    'Dim OthersCon As SqlConnection
    ''Create DataTable for results
    'OthersCon = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
    'OthersCon.Open()
    'Dim TargetConnection As SqlConnection
    'TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
    'TargetConnection.Open()
    'db2021Connection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2022;Trusted_Connection=True;")
    'db2021Connection.Open()
    ''Dim Delquery As String = "DELETE FROM CommonNumbers"
    ''OthersCommand = New SqlCommand(Delquery, OthersConnection)
    ''OthersCommand.ExecuteNonQuery()
    'Dim CreateTablesQuery As String
    'Dim TempTableName As String
    'Dim QueryInsert As String
    'Dim queryString1 As String
    'Dim CreateTbCommand As SqlCommand

    'Dim ConnectionString As String
    'Dim o As OleDb.OleDbConnection
    'Dim queryString As String
    'Dim InsertCommand As OleDb.OleDbCommand
    'Dim tblCommonNumbers As String = "[CommonNumbers] nvarchar(255) null,"

    'Dim DtColName As New Data.DataTable
    'DtColName.Clear()
    'DtColName.Columns.Add("ColNames", GetType(System.String))
    'For j As Integer = 0 To TotalNumberOfCDRs - 1
    '    OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))

    '    TempTableName = "_" & OnlyFileNames(j)
    '    TempTableName = TempTableName
    '    DtColName.Rows.Add(TempTableName)
    '    Try
    '        queryString1 = "Drop Table IF EXISTS [" & TempTableName & "]"
    '        CreateTbCommand = New SqlCommand(queryString1, OthersConnection)
    '        CreateTbCommand.ExecuteNonQuery()
    '    Catch ex03 As Exception
    '    End Try
    '    CreateTablesQuery = "CREATE TABLE [" & TempTableName & "] (PhoneNumber nvarchar(255) null)"
    '    tblCommonNumbers = tblCommonNumbers & "[" & TempTableName & "] nvarchar(255) null,"
    '    OthersCommand = New SqlCommand(CreateTablesQuery, OthersConnection)
    '    Try
    '        OthersCommand.ExecuteNonQuery()
    '        'TargetConnection.Close()
    '    Catch ex02 As Exception
    '        MsgBox("creating table", MsgBoxStyle.OkOnly)
    '        'TargetConnection.Close()
    '    End Try
    '    'CreateTablesQuery = "CREATE CLUSTERED INDEX myIdx ON  [" & TempTableName & "](PhoneNumber)"
    '    'OthersCommand = New SqlCommand(CreateTablesQuery, OthersConnection)
    '    'Try
    '    '    OthersCommand.ExecuteNonQuery()
    '    'Catch ex02 As Exception
    '    '    MsgBox("creating table", MsgBoxStyle.OkOnly)
    '    'End Try
    '    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
    '    o = New OleDb.OleDbConnection(ConnectionString)
    '    o.Open()
    '    'populate table with data with 92
    '    queryString = "select Party from (select [" & ColumnName1 & "] as Party from [" & SheetName & "] WHERE Isnumeric(" & ColumnName1 & ") AND LEN(" & ColumnName1 & ")>=10) " + _
    '                "UNION ALL (select [" & ColumnName2 & "] as Party from [" & SheetName & "] WHERE Isnumeric(" & ColumnName2 & ") AND LEN(" & ColumnName2 & ")>=10)"
    '    InsertCommand = New OleDb.OleDbCommand(queryString, o)
    '    Dim InsertReader As OleDb.OleDbDataReader
    '    Try
    '        InsertReader = InsertCommand.ExecuteReader()
    '    Catch ex02 As Exception
    '        MsgBox("insertion error table", MsgBoxStyle.OkOnly)
    '    End Try
    '    Dim bcCopy As New SqlBulkCopy(TargetConnection)
    '    'TargetConnection.Open()
    '    bcCopy.BatchSize = 100000
    '    bcCopy.BulkCopyTimeout = 0
    '    bcCopy.DestinationTableName = "[" & TempTableName & "]"
    '    bcCopy.WriteToServer(InsertReader)
    '    InsertReader.Close()
    '    'strConSrc.Close()
    '    'TargetConnection.Close()
    '    'End While
    '    'InsertReader.Close()
    '    o.Close()
    '    TransferedFiles = TransferedFiles + 1
    '    frm_Spy_Tech.prgbar_Common_Links.Value = frm_Spy_Tech.prgbar_Common_Links.Value + 1
    '    'frm_Spy_Tech.prgbar_Common_Links.Refresh()
    '    frm_Spy_Tech.lbCommonNos.Text = "Chaching the data " & TransferedFiles & " of " & TotalNumberOfCDRs
    '    frm_Spy_Tech.Refresh()
    'Next
    'End Function
    Public Function csvTOxlsx(ByVal csvFile As String) As String
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        xlWorkBook = xlApp.Workbooks.Open(csvFile)
        xlWorkSheet = xlWorkBook.ActiveSheet
        Dim XlFile As String = Path.ChangeExtension(csvFile, ".xlsx")
        xlApp.DisplayAlerts = False
        xlWorkSheet.SaveAs(XlFile, Excel.XlFileFormat.xlOpenXMLWorkbook)
        Try
            xlWorkBook.Close()
            releaseObject(xlWorkBook)
        Catch ex As Exception

        End Try
        Try
            xlApp.Quit()
            releaseObject(xlApp)
        Catch ex As Exception

        End Try

        Return XlFile
    End Function
    Public Function csvFormat(ByVal csvFile As String)
        Using parser As New TextFieldParser(csvFile)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters("|")

            Using writer As New StreamWriter(csvFile, True)
                While Not parser.EndOfData
                    Dim fields As String() = parser.ReadFields()

                    For i As Integer = 0 To fields.Length - 1
                        If IsZipCode(fields(i)) Then
                            fields(i) = fields(i).PadLeft(5, "0"c)
                        End If
                    Next

                    writer.WriteLine(String.Join("|", fields))
                End While
            End Using
        End Using
    End Function
    Function IsZipCode(ByVal zip As String) As Boolean
        Return (zip.Length = 5 AndAlso IsNumeric(zip))
    End Function
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Module
