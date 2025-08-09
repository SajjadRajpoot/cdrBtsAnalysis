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

Public Class Form1

    Function ChangeTimeFormat(ByVal formatCode As String)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Try


            xlWorkbook = xlWorkbooks.Open(txt_BTS_Path.Text)
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "bts" Then
                    xlSheet = xlWorkbook.Worksheets("bts")
                    IsSheetRenamed = True
                    Exit For
                End If
            Next xlSheet
            If IsSheetRenamed = False Then
                MsgBox("Please rename the sheet as 'bts' and browse again", MsgBoxStyle.OkOnly)
                GoTo Incomplete
            End If
            ' xlSheet = xlWorkbook.Worksheets("bts")
            Dim IsFormatChanged As Boolean = False
            Dim ColNumber As Integer = 1
            Dim ColName As String
            Dim ColRange As String
            While IsFormatChanged = False
                If xlSheet.Cells(1, ColNumber).Value = "Time" Then
                    ColName = Chr(65 + ColNumber - 1)
                    ColRange = ColName & ":" & ColName
                    xlSheet.Range(ColRange).NumberFormat = formatCode
                    IsFormatChanged = True
                End If
                ColNumber = ColNumber + 1
                If ColNumber > 8 Then
                    MsgBox("Please  rename as 'Time' of time column and browse again", MsgBoxStyle.OkOnly)
                    GoTo Incomplete
                End If
            End While
            IsFormatChanged = False
            ColNumber = 1
            While IsFormatChanged = False
                If xlSheet.Cells(1, ColNumber).Value = "a" Then
                    ColName = Chr(65 + ColNumber - 1)
                    ColRange = ColName & ":" & ColName
                    xlSheet.Range(ColRange).NumberFormat = "0"
                    IsFormatChanged = True
                End If
                ColNumber = ColNumber + 1
                If ColNumber > 8 Then
                    MsgBox("Please set then name 'a' of A Party column and browse again", MsgBoxStyle.OkOnly)
                    GoTo Incomplete
                End If
            End While
Incomplete:
            xlApp.DisplayAlerts = False
            ' xlSheet.SaveAs(txt_BTS_Path.Text)
            ' xlApp.DisplayAlerts = True
            xlWorkbook.SaveAs(txt_BTS_Path.Text, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkbooks = Nothing
            xlWorkbook.Close(True, misValue, misValue)
            xlApp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) : xlWorkbook = Nothing
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbooks) : xlWorkbooks = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet) : xlSheet = Nothing
        Catch ex As Exception
            MsgBox("An Error occured: " & ex.Message)
            xlWorkbooks = Nothing
            xlWorkbook.Close()
            xlApp.Quit()
        End Try
    End Function
    Dim StopAnalyze As Boolean = False
    Sub createTblCNIC(ByVal TableName As String)
        Try
            'Call DropTempTable()


            'If TargetConnection.State <> ConnectionState.Open Then
            '    TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            '    TargetConnection.Open()
            'End If
            Call ConnectionOpen()
            'QueryString = "IF OBJECT_ID('tempdb.#IMEITable') IS NOT NULL DROP TABLE #IMEITable"
            'OthersCommand = New SqlCommand(QueryString, TargetConnection)
            'OthersCommand.ExecuteNonQuery()
            'QueryString = "CREATE TABLE #IMEITable(IMEI varchar(16))"

            'QueryString = "CREATE TABLE " & TableName & "(IMEI varchar(16))"
            'CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            'CreateTbCommand.ExecuteNonQuery()
            'TargetConnection.Close()
            'TargetConnection.Dispose()
            'CreateTbCommand.Dispose()

            QueryString = "CREATE TABLE [" & TableName & "] (PhoneNumber varchar(16),CNIC text)"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
            TargetConnection.Dispose()
            CreateTbCommand.Dispose()

            'TargetConnection.Close()
            'TargetConnection.Close()
        Catch ex As Exception

            'If TargetConnection.State <> ConnectionState.Closed Then
            '    TargetConnection.Close()
            'End If
            'TargetConnection.Close()
        End Try
    End Sub
    Dim CNICQuery As String
    Dim CNICOthersCommand As SqlCommand
    Dim CNICReader As SqlDataReader
    Function GetRptNumberCNIC(ByVal phnNumber As String)
        Call ConnectionOpen()
        CNICQuery = "Select CNIC from tblCNIC where PhoneNumber = '" & phnNumber.ToString & "'"
        CNICOthersCommand = New SqlCommand(CNICQuery, TargetConnection)
        CNICReader = CNICOthersCommand.ExecuteReader()
        GprCNIC = Nothing
        While CNICReader.Read()
            If IsDBNull(CNICReader(0)) = False Then
                GprCNIC = CNICReader(0).ToString
            Else
                GprCNIC = Nothing
            End If
        End While
        CNICReader.Close()
        TargetConnection.Close()
        CNICOthersCommand.Dispose()
        CNICReader.Close()
        CNICQuery = Nothing
    End Function
    Function MakeCopyOfExcel(ByVal pathOfFile As String) As String
        Dim CopyOfFilePath As String = pathOfFile.Insert(pathOfFile.IndexOf("."), "1")

        If System.IO.File.Exists(pathOfFile) = True Then
            My.Computer.FileSystem.CopyFile(pathOfFile, CopyOfFilePath, True)
        End If
        Return CopyOfFilePath
    End Function


    Sub CreateTblOutOfLimit(ByVal tblName As String)
        Try
            Call ConnectionOpenDelDupli()
            QueryString = "IF OBJECT_ID('dbo." & tblName & "') IS NOT NULL DROP TABLE " & tblName & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnectionDelDupli)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnectionDelDupli.Close()
        Catch ex As Exception
            If TargetConnectionDelDupli.State <> ConnectionState.Closed Then
                TargetConnectionDelDupli.Close()
            End If
        End Try
        Call ConnectionOpenDelDupli()
        Try
            QueryString = "CREATE TABLE " & tblName & "(a float)"
            'QueryString = "CREATE TABLE " & tblName & "(" & ColumnsHeaders & ")"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnectionDelDupli)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnectionDelDupli.Close()
            TargetConnectionDelDupli.Dispose()
            CreateTbCommand.Dispose()
        Catch ex As Exception

        End Try
    End Sub
    
    Sub InsertaPartyOutOfLimit(ByVal PhoneNumber As String)
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO bParty VALUES (' & PhoneNumber & ')"
        '" & CNIC & "'"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        InsertIMEICommand.ExecuteNonQuery()

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub XL_to_SQL()
        Dim SheetName As String = "bts$"
        Dim queryString As String = "Select * from [" & SheetName & "]" 'Order By '" & A_party & "'"
        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & txt_BTS_Path.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        Dim FilterQueryString As String
        Dim oreader As OleDb.OleDbDataReader
        Dim ocmd1 As New OleDb.OleDbCommand(queryString, o)
        Dim oreader1 As OleDb.OleDbDataReader
        oreader1 = ocmd1.ExecuteReader()
        Call CreateNewExcelFile()
        Dim ColNumber As Integer = 1
        Dim ColName As String
        Dim ColRange As String
        Dim FieldsNames As String
        Dim totalfiels As Integer = oreader1.FieldCount
        If chkCNIC.Checked = False Then
            For i As Integer = 1 To totalfiels
                newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 1).ToString
            Next
        Else
            For i As Integer = 1 To totalfiels + 2
                If i = 1 Then
                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 1).ToString
                    FieldsNames = "[" & oreader1.GetName(i - 1).ToString & "] varchar(100)"
                ElseIf i = 2 Then
                    newXlWorkSheet.Cells(1, i) = "CNIC_a"
                    FieldsNames = FieldsNames & ", [CNIC_a] varchar(255)"
                ElseIf i = 3 Then
                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 2).ToString
                    FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 2).ToString & "] varchar(100)"
                ElseIf i = 4 Then
                    newXlWorkSheet.Cells(1, i) = "CNIC_b"
                    FieldsNames = FieldsNames & ", [CNIC_b] varchar(255)"
                ElseIf i >= 5 Then
                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 3).ToString
                    If oreader1.GetName(i - 3).ToString = "Time" Then
                        FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 3).ToString & "] float"
                    Else
                        FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 3).ToString & "] varchar(255)"
                    End If

                End If
            Next
            Call CreateTblLimimtedActivity("CompleteXLFile", FieldsNames)
        End If
        lbNoOfRecords.Text = "Initializing...   Step3"
        lbNoOfRecords.Refresh()
        'Cutting within limit records
        Dim PrevousNumber As String = Nothing
        Dim currentNumber As Double
        Dim fieldName As String = oreader1.GetName(0).ToString
        Dim IndexOfRow As Integer = 2
        Dim IsCulprit As Boolean = False
        Dim InsertingValues As String
        Dim TotalRecords As Integer = 0
        Dim TimeToDouble As Double
        Dim MytimeDate As Date
        While oreader1.Read
            InsertingValues = Nothing
            For i As Integer = 1 To totalfiels + 2
                If i = 1 Then
                    If IsDBNull(oreader1(i - 1)) = False Then
                        InsertingValues = "'" & oreader1(i - 1) & "'"
                    Else
                        InsertingValues = "'NULL'"
                    End If

                ElseIf i = 2 Then
                    InsertingValues = InsertingValues & ", " & "'NULL'"
                ElseIf i = 3 Then
                    If IsDBNull(oreader1(i - 2)) = False Then
                        InsertingValues = InsertingValues & ", " & "'" & oreader1(i - 2) & "'"
                    Else
                        InsertingValues = InsertingValues & ", 'NULL'"
                    End If

                ElseIf i = 4 Then

                    InsertingValues = InsertingValues & ", " & "'NULL'"

                ElseIf i >= 5 Then
                    If oreader1.GetName(i - 3).ToString = "Time" Then
                        If IsDBNull(oreader1(i - 3)) = False Then
                            MytimeDate = oreader1(i - 3)
                            TimeToDouble = MytimeDate.ToOADate
                            InsertingValues = InsertingValues & ", '" & TimeToDouble & "'"

                        Else
                            InsertingValues = InsertingValues & ", 'NULL'"
                        End If
                    Else
                        If IsDBNull(oreader1(i - 3)) = False Then
                            InsertingValues = InsertingValues & ", '" & oreader1(i - 3) & "'"
                        Else
                            InsertingValues = InsertingValues & ", 'NULL'"
                        End If
                    End If

                   
                End If
            Next
            Call InsertAnalyzedToSQL("CompleteXLFile", InsertingValues)
            TotalRecords = TotalRecords + 1
            lbBT_CRP_Checked.Text = TotalRecords
            lbBT_CRP_Checked.Refresh()
        End While
        lbNoOfRecords.Text = "Initializing...   Step4"
        lbNoOfRecords.Refresh()

    End Sub
    Dim InLimitSqlConn As SqlConnection
    Dim InLimitSqlCommand As SqlCommand
    Dim InlimitSqlReader As SqlDataReader

    Sub ProblematicXLFile()
        Dim ColumnName As String = "Time"
        Dim SheetName As String = "bts$"
        Dim A_party As String = "a"
        Dim intialTime As String = dtp_Initial_Time.Value.ToOADate
        Dim endTime As String = dtp_Ending_Time.Value.ToOADate
        Dim misValue As Object = System.Reflection.Missing.Value
        If chkCNIC.Checked = True Then
            Call DropTempTable("tblCNIC")
            Call createTblCNIC("tblCNIC")
        End If

        Dim queryString As String = "Select * from CompleteXLFile where [" & ColumnName & "] between '" & intialTime & "' AND '" & endTime & "'" 'Order By '" & A_party & "'"
        Call InlimitConnection()
        InLimitSqlCommand = New SqlCommand(queryString, InLimitSqlConn)
        InlimitSqlReader = InLimitSqlCommand.ExecuteReader
        Call CreateNewExcelFile()
        Dim ColNumber As Integer = 1
        Dim ColName As String
        Dim ColRange As String
        Dim FieldsNames As String
        Dim totalfiels As Integer = InlimitSqlReader.FieldCount
        If chkCNIC.Checked = False Then
            For i As Integer = 0 To totalfiels - 1
                newXlWorkSheet.Cells(1, i + 1) = InlimitSqlReader.GetName(i).ToString
            Next
        Else
            For i As Integer = 0 To totalfiels - 1
                If i = 0 Then
                    newXlWorkSheet.Cells(1, i + 1) = InlimitSqlReader.GetName(i).ToString
                    FieldsNames = "[" & InlimitSqlReader.GetName(i).ToString & "] varchar(100)"
                ElseIf i = 1 Then
                    newXlWorkSheet.Cells(1, i + 1) = "CNIC_a"
                    FieldsNames = FieldsNames & ", [CNIC_a] varchar(255)"
                ElseIf i = 2 Then
                    newXlWorkSheet.Cells(1, i + 1) = InlimitSqlReader.GetName(i).ToString
                    FieldsNames = FieldsNames & ", [" & InlimitSqlReader.GetName(i).ToString & "] varchar(100)"
                ElseIf i = 3 Then
                    newXlWorkSheet.Cells(1, i + 1) = "CNIC_b"
                    FieldsNames = FieldsNames & ", [CNIC_b] varchar(255)"

                ElseIf i >= 4 Then
                    newXlWorkSheet.Cells(1, i + 1) = InlimitSqlReader.GetName(i).ToString
                    FieldsNames = FieldsNames & ", [" & InlimitSqlReader.GetName(i).ToString & "] varchar(255)"
                End If
            Next
            Call CreateTblLimimtedActivity("TblWithinLimit", FieldsNames)
        End If
        lbNoOfRecords.Text = "Initializing...   Step3"
        lbNoOfRecords.Refresh()
        'Cutting within limit records
        Dim PrevousNumber As String = Nothing
        Dim currentNumber As Double
        Dim fieldName As String = InlimitSqlReader.GetName(0).ToString
        Dim IndexOfRow As Integer = 2
        Dim IsCulprit As Boolean = False
        Dim InsertingValues As String
        While InlimitSqlReader.Read
            InsertingValues = Nothing
            For i As Integer = 0 To totalfiels - 1
                If i = 0 Then
                    If IsDBNull(InlimitSqlReader(i)) = False Then
                        InsertingValues = "'" & InlimitSqlReader(i) & "'"
                    Else
                        InsertingValues = "'NULL'"
                    End If

                ElseIf i = 1 Then
                    InsertingValues = InsertingValues & ", " & "'NULL'"
                ElseIf i = 2 Then
                    If IsDBNull(InlimitSqlReader(i)) = False Then
                        InsertingValues = InsertingValues & ", " & "'" & InlimitSqlReader(i) & "'"
                    Else
                        InsertingValues = InsertingValues & ", 'NULL'"
                    End If

                ElseIf i = 3 Then

                    InsertingValues = InsertingValues & ", " & "'NULL'"

                ElseIf i >= 4 Then
                    If IsDBNull(InlimitSqlReader(i)) = False Then
                        InsertingValues = InsertingValues & ", '" & InlimitSqlReader(i) & "'"
                    Else
                        InsertingValues = InsertingValues & ", 'NULL'"
                    End If
                End If
            Next
            Call InsertAnalyzedToSQL("TblWithinLimit", InsertingValues)
        End While
        lbNoOfRecords.Text = "Initializing...   Step4"
        lbNoOfRecords.Refresh()
        InlimitSqlReader.Close()
        InLimitSqlConn.Close()
        InLimitSqlCommand.Dispose()
        'Cutting out of limit records
        Call CreateTblOutOfLimit("OutOfLimit")
        Dim FilterQueryString As String = "Select Distinct(a) from CompleteXLFile where ([" & ColumnName & "] < '" & intialTime & "' OR [" & ColumnName & "] > '" & endTime & "')"
        Call OutlimitConnection()

        OutLimitSqlCommand = New SqlCommand(FilterQueryString, OutLimitSqlConn)
        OutLimitSqlReader = OutLimitSqlCommand.ExecuteReader
        Call OutOfLimmitConn()
        Dim InserOutOfLimitCommand As SqlCommand
        Dim SqlOutOfLimitReder As SqlDataReader
        Dim currentPhnNumber As Double
        While OutLimitSqlReader.Read
            If IsDBNull(OutLimitSqlReader(A_party)) = False Then
                currentPhnNumber = OutLimitSqlReader(A_party)
                InsertQuery = "INSERT INTO OutOfLimit VALUES ('" & currentPhnNumber & "')"
                InserOutOfLimitCommand = New SqlCommand(InsertQuery, OutOfLimitConnection)
                InserOutOfLimitCommand.ExecuteNonQuery()
            End If
        End While
        OutLimitSqlReader.Close()
        OutLimitSqlConn.Close()
        OutLimitSqlCommand.Dispose()
        OutOfLimitConnection.Close()
        OutOfLimitConnection.Dispose()
        InserOutOfLimitCommand.Dispose()

        lbNoOfRecords.Text = "Initializing...   Step5"
        lbNoOfRecords.Refresh()
        Dim CheckedTotal As Long = 0
        queryString = "Select Distinct(a) from TblWithinLimit"
        Call ConnectionOpen()
        Dim WithinLimitCommand As SqlCommand
        Dim WithinLimitReader As SqlDataReader
        WithinLimitCommand = New SqlCommand(queryString, TargetConnection)
        WithinLimitReader = WithinLimitCommand.ExecuteReader()
        While WithinLimitReader.Read
            If IsDBNull(WithinLimitReader(0)) = False Then
                If isOutofLimit(WithinLimitReader(0)) = True Then
                    PurifyingData(WithinLimitReader(0))
                End If
            End If
            CheckedTotal = CheckedTotal + 1
            lbBT_CRP_Checked.Text = CheckedTotal
            lbBT_CRP_Checked.Refresh()
        End While
        WithinLimitReader.Close()
        TargetConnection.Close()
        TargetConnection.Dispose()
        WithinLimitCommand.Dispose()
        lbBT_CRP_Checked.Text = "Croped..."
        lbBT_CRP_Checked.Refresh()
        '  UpdatingCNICinBTS("a")
        ' UpdatingCNICinBTS("b")
        'Inserting into Excel File
        lbNoOfRecords.Text = ""
        lbNoOfRecords.Refresh()
        lbNoOfRecords.Text = "0"
        lbNoOfRecords.Refresh()
        queryString = "Select * from TblWithinLimit"
        Call OutOfLimmitConn()
        WithinLimitCommand = New SqlCommand(queryString, OutOfLimitConnection)
        WithinLimitReader = WithinLimitCommand.ExecuteReader()
        While WithinLimitReader.Read
            If chkCNIC.Checked = False Then
                For j As Integer = 1 To totalfiels
                    newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                    IsCulprit = True
                Next
                IndexOfRow = IndexOfRow + 1
                lbNoOfRecords.Text = IndexOfRow - 2
            Else
                GprCNIC = Nothing
                For j As Integer = 1 To totalfiels + 2
                    If j = 1 Then

                        If IsDBNull(WithinLimitReader(j - 1)) = False And WithinLimitReader(j - 1) <> "NULL" Then
                            newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                            IsCNIC_Need = True
                            Call GetRptNumberCNIC(WithinLimitReader(j - 1).ToString)
                            If GprCNIC = Nothing Then
                                Call FindNumberDB2021(WithinLimitReader(j - 1))
                                '' Call FindPhoneNumber(WithinLimitReader(j - 1))
                                If GprCNIC = Nothing Then
                                    GprCNIC = ""
                                End If
                                Call InsertCNIC_CropBTS(WithinLimitReader(j - 1).ToString, GprCNIC.ToString, "tblCNIC")
                            End If
                            IsCNIC_Need = False
                        End If
                    ElseIf j = 2 Then
                        If GprCNIC = Nothing Then
                            GprCNIC = ""
                        End If
                        newXlWorkSheet.Cells(IndexOfRow, j) = GprCNIC.ToString

                        GprCNIC = Nothing
                    ElseIf j = 3 Then
                        If IsDBNull(WithinLimitReader(j - 1)) = False And WithinLimitReader(j - 1) <> "NULL" Then
                            newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                            IsCNIC_Need = True
                            Call GetRptNumberCNIC(WithinLimitReader(j - 1).ToString)
                            If GprCNIC = Nothing Then
                                Call FindNumberDB2021(WithinLimitReader(j - 1))
                                ''Call FindPhoneNumber(WithinLimitReader(j - 1))
                                If GprCNIC = Nothing Then
                                    GprCNIC = ""
                                End If
                                Call InsertCNIC_CropBTS(WithinLimitReader(j - 1).ToString, GprCNIC.ToString, "tblCNIC")
                            End If
                            IsCNIC_Need = False
                        End If
                    ElseIf j = 4 Then
                        If GprCNIC = Nothing Then
                            GprCNIC = ""
                        End If
                        newXlWorkSheet.Cells(IndexOfRow, j) = GprCNIC.ToString
                        GprCNIC = Nothing
                    ElseIf j >= 5 Then
                        newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                    End If

                    IsCulprit = True
                Next
                IndexOfRow = IndexOfRow + 1
                lbNoOfRecords.Text = IndexOfRow - 2
            End If
        End While
        WithinLimitReader.Close()
        OutOfLimitConnection.Close()
        WithinLimitCommand.Dispose()
        lbNoOfRecords.Text = IndexOfRow - 2 & " ... Finalizing ...Step1"
        lbNoOfRecords.Refresh()
        For i As Integer = 1 To totalfiels
            'newXlWorkSheet.Cells(1, i).value()
            If newXlWorkSheet.Cells(1, i).value = "a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "CNIC_a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "CNIC_b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "Time" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "hh:mm:ss AM/PM"
            ElseIf newXlWorkSheet.Cells(1, i).value = "Date" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "dd-mm-yy"
            ElseIf newXlWorkSheet.Cells(1, i).value = "b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "IMEI" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            End If


            ColNumber = ColNumber + 1
        Next
        Try
            newXlWorkSheet.Columns.AutoFit()
        Catch ex As Exception

        End Try
        lbNoOfRecords.Text = IndexOfRow - 2 & " ... Finalizing ...Step2"
        lbNoOfRecords.Refresh()
        Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(txt_BTS_Path.Text)
        Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(txt_BTS_Path.Text)
        Dim JustExtention As String = System.IO.Path.GetExtension(txt_BTS_Path.Text)
        newXlApp.DisplayAlerts = False
        Try
            If chkCNIC.Checked = False Then
                newXlWorkSheet.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention)
            Else
                newXlWorkSheet.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)CNIC" & JustExtention)
            End If
        Catch ex As Exception
            If chkCNIC.Checked = False Then
                MsgBox("Please close the file" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.Information)
            Else
                MsgBox("Please close the file" & JustFileName & "(Analyzed)CNIC" & JustExtention, MsgBoxStyle.Information)
            End If
        End Try

        'newXlWorkbook.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        newXlWorkbook.Close()
        newXlWorkbooks = Nothing
        newXlApp.Quit()
        newXlApp = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlApp) : newXlApp = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkbook) : newXlWorkbook = Nothing
        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkbooks) : newXlWorkbooks = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkSheet) : newXlWorkSheet = Nothing
        ' MsgBox("okay" & NumberOfRows, MsgBoxStyle.OkOnly)
        'oreader1.Close()
        ' TargetConnection.Close()
        'TargetConnection.Dispose()
        OutOfLimitConnection.Close()
        OutOfLimitConnection.Dispose()
        ' oreader.Close()
        ' ocmd1.Dispose()
        'If o.State <> ConnectionState.Closed Then
        '    o.Close()
        '    o.Dispose()

        'End If

        'o = Nothing
        Call ChangeTimeFormat("h:mm:ss AM/PM")
        lbNoOfRecords.Text = IndexOfRow - 2 & " ... Finalized"
        lbNoOfRecords.Refresh()
        Me.Cursor = Cursors.Default
        Button1.Text = "Close"
        ' MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.OkOnly)
        If chkCNIC.Checked = False Then
            lbfilepath.Text = DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention
        Else
            lbfilepath.Text = DirectoryPath & "\" & JustFileName & "(Analyzed)CNIC" & JustExtention
        End If
        If cbFind_b_in_a.Checked = True Then
            ' Call Find_b_in_a(lbfilepath.Text)
            NumberOfFiles = 1
            AllCommonFileNames(0) = lbfilepath.Text
            Call FindGroupsOfOne()
            MsgBox("Both files 'Analyzed and Analyzed Groups' have been created", MsgBoxStyle.Information)
            ' MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.OkOnly)
        Else
            MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.Information)
        End If
    End Sub
    Sub CutWithinLimit()
        Dim ColumnName As String = "Time"
        Dim SheetName As String = "bts$"
        Dim A_party As String = "a"
        Dim intialTime As String = dtp_Initial_Time.Value.ToOADate
        Dim endTime As String = dtp_Ending_Time.Value.ToOADate
        Dim misValue As Object = System.Reflection.Missing.Value
        If chkCNIC.Checked = True Then
            Call DropTempTable("tblCNIC")
            Call createTblCNIC("tblCNIC")
        End If

        Dim queryString As String = "Select * from [" & SheetName & "] where [" & ColumnName & "] between '" & intialTime & "' AND '" & endTime & "'" 'Order By '" & A_party & "'"
        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & txt_BTS_Path.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        Dim FilterQueryString As String
        Dim oreader As OleDb.OleDbDataReader
        Dim ocmd1 As New OleDb.OleDbCommand(queryString, o)
        Dim oreader1 As OleDb.OleDbDataReader
        oreader1 = ocmd1.ExecuteReader()
        Call CreateNewExcelFile()
        Dim ColNumber As Integer = 1
        Dim ColName As String
        Dim ColRange As String
        Dim FieldsNames As String
        Dim totalfiels As Integer = oreader1.FieldCount
        If chkCNIC.Checked = False Then
            For i As Integer = 1 To totalfiels
                newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 1).ToString
            Next
        Else
            For i As Integer = 1 To totalfiels + 2
                If i = 1 Then
                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 1).ToString
                    FieldsNames = "[" & oreader1.GetName(i - 1).ToString & "] varchar(100)"
                ElseIf i = 2 Then
                    newXlWorkSheet.Cells(1, i) = "CNIC_a"
                    FieldsNames = FieldsNames & ", [CNIC_a] varchar(255)"
                ElseIf i = 3 Then
                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 2).ToString
                    FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 2).ToString & "] varchar(100)"
                ElseIf i = 4 Then
                    newXlWorkSheet.Cells(1, i) = "CNIC_b"
                    FieldsNames = FieldsNames & ", [CNIC_b] varchar(255)"

                ElseIf i >= 5 Then
                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 3).ToString
                    FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 3).ToString & "] varchar(255)"
                End If
            Next
            Call CreateTblLimimtedActivity("TblWithinLimit", FieldsNames)
        End If
        lbNoOfRecords.Text = "Initializing...   Step3"
        lbNoOfRecords.Refresh()
        'Cutting within limit records
        Dim PrevousNumber As String = Nothing
        Dim currentNumber As Double
        Dim fieldName As String = oreader1.GetName(0).ToString
        Dim IndexOfRow As Integer = 2
        Dim IsCulprit As Boolean = False
        Dim InsertingValues As String
        While oreader1.Read
            InsertingValues = Nothing
            For i As Integer = 1 To totalfiels + 2
                If i = 1 Then
                    If IsDBNull(oreader1(i - 1)) = False Then
                        InsertingValues = "'" & oreader1(i - 1) & "'"
                    Else
                        InsertingValues = "'NULL'"
                    End If

                ElseIf i = 2 Then
                    InsertingValues = InsertingValues & ", " & "'NULL'"
                ElseIf i = 3 Then
                    If IsDBNull(oreader1(i - 2)) = False Then
                        InsertingValues = InsertingValues & ", " & "'" & oreader1(i - 2) & "'"
                    Else
                        InsertingValues = InsertingValues & ", 'NULL'"
                    End If

                ElseIf i = 4 Then

                    InsertingValues = InsertingValues & ", " & "'NULL'"

                ElseIf i >= 5 Then
                    If IsDBNull(oreader1(i - 3)) = False Then
                        InsertingValues = InsertingValues & ", '" & oreader1(i - 3) & "'"
                    Else
                        InsertingValues = InsertingValues & ", 'NULL'"
                    End If
                End If
            Next
            Call InsertAnalyzedToSQL("TblWithinLimit", InsertingValues)
        End While
        lbNoOfRecords.Text = "Initializing...   Step4"
        lbNoOfRecords.Refresh()
        'Cutting out of limit records
        Call CreateTblOutOfLimit("OutOfLimit")
        FilterQueryString = "Select Distinct(a) from [" & SheetName & "] where ([" & ColumnName & "] < '" & intialTime & "' OR [" & ColumnName & "] > '" & endTime & "')"
        Dim Ocmd As New OleDb.OleDbCommand(FilterQueryString, o)
        oreader = Ocmd.ExecuteReader()
        Call OutOfLimmitConn()
        Dim InserOutOfLimitCommand As SqlCommand
        Dim SqlOutOfLimitReder As SqlDataReader
        Dim currentPhnNumber As Double
        While oreader.Read
            If IsDBNull(oreader(A_party)) = False Then
                currentPhnNumber = oreader(A_party)
                InsertQuery = "INSERT INTO OutOfLimit VALUES ('" & currentPhnNumber & "')"
                InserOutOfLimitCommand = New SqlCommand(InsertQuery, OutOfLimitConnection)
                InserOutOfLimitCommand.ExecuteNonQuery()
            End If
        End While
        oreader.Close()
        OutOfLimitConnection.Close()
        OutOfLimitConnection.Dispose()
        InserOutOfLimitCommand.Dispose()

        lbNoOfRecords.Text = "Initializing...   Step5"
        lbNoOfRecords.Refresh()
        Dim CheckedTotal As Long = 0
        queryString = "Select Distinct(a) from TblWithinLimit"
        Call ConnectionOpen()
        Dim WithinLimitCommand As SqlCommand
        Dim WithinLimitReader As SqlDataReader
        WithinLimitCommand = New SqlCommand(queryString, TargetConnection)
        WithinLimitReader = WithinLimitCommand.ExecuteReader()
        While WithinLimitReader.Read
            If IsDBNull(WithinLimitReader(0)) = False Then
                If isOutofLimit(WithinLimitReader(0)) = True Then
                    PurifyingData(WithinLimitReader(0))
                End If
            End If
            CheckedTotal = CheckedTotal + 1
            lbBT_CRP_Checked.Text = CheckedTotal
            lbBT_CRP_Checked.Refresh()
        End While
        WithinLimitReader.Close()
        TargetConnection.Close()
        TargetConnection.Dispose()
        WithinLimitCommand.Dispose()
        lbBT_CRP_Checked.Text = "Croped..."
        lbBT_CRP_Checked.Refresh()
        '  UpdatingCNICinBTS("a")
        ' UpdatingCNICinBTS("b")
        'Inserting into Excel File
        lbNoOfRecords.Text = ""
        lbNoOfRecords.Refresh()
        lbNoOfRecords.Text = "0"
        lbNoOfRecords.Refresh()
        queryString = "Select * from TblWithinLimit"
        Call OutOfLimmitConn()
        WithinLimitCommand = New SqlCommand(queryString, OutOfLimitConnection)
        WithinLimitReader = WithinLimitCommand.ExecuteReader()
        While WithinLimitReader.Read
            If chkCNIC.Checked = False Then
                For j As Integer = 1 To totalfiels
                    newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                    IsCulprit = True
                Next
                IndexOfRow = IndexOfRow + 1
                lbNoOfRecords.Text = IndexOfRow - 2
            Else
                GprCNIC = Nothing
                For j As Integer = 1 To totalfiels + 2
                    If j = 1 Then

                        If IsDBNull(WithinLimitReader(j - 1)) = False And WithinLimitReader(j - 1) <> "NULL" Then
                            newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                            IsCNIC_Need = True
                            Call GetRptNumberCNIC(WithinLimitReader(j - 1).ToString)
                            If GprCNIC = Nothing Then
                                Call FindNumberDB2021(WithinLimitReader(j - 1))
                                '' Call FindPhoneNumber(WithinLimitReader(j - 1))
                                If GprCNIC = Nothing Then
                                    GprCNIC = ""
                                End If
                                Call InsertCNIC_CropBTS(WithinLimitReader(j - 1).ToString, GprCNIC.ToString, "tblCNIC")
                            End If
                            IsCNIC_Need = False
                        End If
                    ElseIf j = 2 Then
                        If GprCNIC = Nothing Then
                            GprCNIC = ""
                        End If
                        newXlWorkSheet.Cells(IndexOfRow, j) = GprCNIC.ToString

                        GprCNIC = Nothing
                    ElseIf j = 3 Then
                        If IsDBNull(WithinLimitReader(j - 1)) = False And WithinLimitReader(j - 1) <> "NULL" Then
                            newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                            IsCNIC_Need = True
                            Call GetRptNumberCNIC(WithinLimitReader(j - 1).ToString)
                            If GprCNIC = Nothing Then
                                Call FindNumberDB2021(WithinLimitReader(j - 1))
                                '' Call FindPhoneNumber(WithinLimitReader(j - 1))
                                If GprCNIC = Nothing Then
                                    GprCNIC = ""
                                End If
                                Call InsertCNIC_CropBTS(WithinLimitReader(j - 1).ToString, GprCNIC.ToString, "tblCNIC")
                            End If
                            IsCNIC_Need = False
                        End If
                    ElseIf j = 4 Then
                        If GprCNIC = Nothing Then
                            GprCNIC = ""
                        End If
                        newXlWorkSheet.Cells(IndexOfRow, j) = GprCNIC.ToString
                        GprCNIC = Nothing
                    ElseIf j >= 5 Then
                        newXlWorkSheet.Cells(IndexOfRow, j) = WithinLimitReader(j - 1)
                    End If

                    IsCulprit = True
                Next
                IndexOfRow = IndexOfRow + 1
                lbNoOfRecords.Text = IndexOfRow - 2
            End If
        End While
        WithinLimitReader.Close()
        OutOfLimitConnection.Close()
        WithinLimitCommand.Dispose()
        lbNoOfRecords.Text = IndexOfRow - 2 & " ... Finalizing ...Step1"
        lbNoOfRecords.Refresh()
        For i As Integer = 1 To totalfiels
            'newXlWorkSheet.Cells(1, i).value()
            If newXlWorkSheet.Cells(1, i).value = "a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "CNIC_a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "CNIC_b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "Time" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "hh:mm:ss AM/PM"
            ElseIf newXlWorkSheet.Cells(1, i).value = "Date" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "dd-mm-yy"
            ElseIf newXlWorkSheet.Cells(1, i).value = "b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, i).value = "IMEI" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            End If


            ColNumber = ColNumber + 1
        Next
        Try
            newXlWorkSheet.Columns.AutoFit()
        Catch ex As Exception

        End Try
        lbNoOfRecords.Text = IndexOfRow - 2 & " ... Finalizing ...Step2"
        lbNoOfRecords.Refresh()
        Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(txt_BTS_Path.Text)
        Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(txt_BTS_Path.Text)
        Dim JustExtention As String = System.IO.Path.GetExtension(txt_BTS_Path.Text)
        newXlApp.DisplayAlerts = False
        Try
            If chkCNIC.Checked = False Then
                newXlWorkSheet.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention)
            Else
                newXlWorkSheet.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)CNIC" & JustExtention)
            End If
        Catch ex As Exception
            If chkCNIC.Checked = False Then
                MsgBox("Please close the file" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.Information)
            Else
                MsgBox("Please close the file" & JustFileName & "(Analyzed)CNIC" & JustExtention, MsgBoxStyle.Information)
            End If
        End Try

        'newXlWorkbook.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        newXlWorkbook.Close()
        newXlWorkbooks = Nothing
        newXlApp.Quit()
        newXlApp = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlApp) : newXlApp = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkbook) : newXlWorkbook = Nothing
        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkbooks) : newXlWorkbooks = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkSheet) : newXlWorkSheet = Nothing
        ' MsgBox("okay" & NumberOfRows, MsgBoxStyle.OkOnly)
        oreader1.Close()
        ' TargetConnection.Close()
        'TargetConnection.Dispose()
        OutOfLimitConnection.Close()
        OutOfLimitConnection.Dispose()
        ' oreader.Close()
        ocmd1.Dispose()
        If o.State <> ConnectionState.Closed Then
            o.Close()
            o.Dispose()

        End If

        'o = Nothing
        Call ChangeTimeFormat("h:mm:ss AM/PM")
        lbNoOfRecords.Text = IndexOfRow - 2 & " ... Finalized"
        lbNoOfRecords.Refresh()
        Me.Cursor = Cursors.Default
        Button1.Text = "Close"
        ' MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.OkOnly)
        If chkCNIC.Checked = False Then
            lbfilepath.Text = DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention
        Else
            lbfilepath.Text = DirectoryPath & "\" & JustFileName & "(Analyzed)CNIC" & JustExtention
        End If
        If cbFind_b_in_a.Checked = True Then
            ' Call Find_b_in_a(lbfilepath.Text)
            NumberOfFiles = 1
            AllCommonFileNames(0) = lbfilepath.Text
            Call FindGroupsOfOne()
            MsgBox("Both files 'Analyzed and Analyzed Groups' have been created", MsgBoxStyle.Information)
            ' MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.OkOnly)
        Else
            MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.Information)
        End If
    End Sub

    Sub UpdatingCNICinBTS(ByVal party As String)
        'Call ConnectionOpen()
        Call OutOfLimmitConn()
        Call ConWithinLimmit()
        Call ConUpdatingInLimit()
        Dim FoundRecords As Long = 0
        Dim queryString As String = "Select Distinct(" + party + ") from TblWithinLimit"
        Dim queryStringUpdating As String
        Dim cmdDistinctAs As SqlCommand
        Dim cmdUpdating As SqlCommand
        Dim DistinctAsReader As SqlDataReader
        cmdDistinctAs = New SqlCommand(queryString, OutOfLimitConnection)
        DistinctAsReader = cmdDistinctAs.ExecuteReader
        While DistinctAsReader.Read
            If IsDBNull(DistinctAsReader(0)) = False Then
                IsCNIC_Need = True
                Call FindNumberDB2021(DistinctAsReader(0))
                ''Call FindPhoneNumber(DistinctAsReader(0))
                If GprCNIC <> Nothing Then
                    queryStringUpdating = "UPDATE TblWithinLimit SET CNIC_" + party + " = '" & GprCNIC & "' Where " + party + " = '" & DistinctAsReader(0) & "'"
                    cmdUpdating = New SqlCommand(queryStringUpdating, ConnectionUpdatingLimit)
                    cmdUpdating.ExecuteNonQuery()
                End If
                IsCNIC_Need = False
            End If
            FoundRecords = FoundRecords + 1
            lbNoOfRecords.Text = FoundRecords
            lbNoOfRecords.Refresh()
        End While
        DistinctAsReader.Close()
        OutOfLimitConnection.Close()
        cmdDistinctAs.Dispose()
        ConnectionUpdatingLimit.Close()
        cmdUpdating.Dispose()
    End Sub
    Function isOutofLimit(ByVal CurrentNumber As String) As Boolean
        Dim OutOfLimit As Boolean = False
        Call OutOfLimmitConn()
        Dim OutofLimitCommand As SqlCommand
        Dim OutofLimitReader As SqlDataReader
        Dim queryString As String = "Select a from OutOfLimit where a =" & "'" & CurrentNumber & "'"
        OutofLimitCommand = New SqlCommand(queryString, OutOfLimitConnection)
        OutofLimitReader = OutofLimitCommand.ExecuteReader()
        While OutofLimitReader.Read
            OutOfLimit = True
            Exit While
        End While
        OutofLimitReader.Close()
        OutOfLimitConnection.Close()
        OutOfLimitConnection.Dispose()
        OutofLimitCommand.Dispose()
        Return OutOfLimit
    End Function
    Sub PurifyingData(ByVal CurrentNumber As String)
        Dim withinLimitConnection As SqlConnection
        withinLimitConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If withinLimitConnection.State = ConnectionState.Closed Then
            withinLimitConnection.Open()
        End If
        Dim queryString As String = "Delete from TblWithinLimit where a = '" & CurrentNumber & "'"
        Dim CMDwithinlimit As SqlCommand
        CMDwithinlimit = New SqlCommand(queryString, withinLimitConnection)
        CMDwithinlimit.ExecuteNonQuery()
        withinLimitConnection.Close()
        withinLimitConnection.Dispose()
        CMDwithinlimit.Dispose()
        queryString = Nothing
    End Sub
    Private Sub btn_CropBTS_Click(sender As System.Object, e As System.EventArgs) Handles btn_CropBTS.Click
        'Button1.Text = "Stop"
        ' Call CreateTblOutOfLimit("OutOfLimit")
        ' Label3.Text = "Initialization"
        Dim CheckedTotal As Long = 0
        ' If chkCNIC.Visible = True Then
        If chkCNIC.Checked = False Then
            Dim Reply As MsgBoxResult
            Reply = MsgBox("Do you want to create file with 'CNIC'", MsgBoxStyle.YesNoCancel)
            If Reply = MsgBoxResult.Yes Then
                chkCNIC.Checked = True
            ElseIf Reply = MsgBoxResult.Cancel Then
                Exit Sub
            ElseIf Reply = MsgBoxResult.No Then
                chkLimitedActivity.Checked = False
            End If
        End If
        '  End If
        If cbProblemFile.Checked = True Then
            Call XL_to_SQL()
            Call ProblematicXLFile()
        Else
            lbNoOfRecords.Text = "Initializing...   Step1"
            lbNoOfRecords.Refresh()
            ' Me.Cursor = Cursors.WaitCursor
            Call ChangeTimeFormat("@")
            lbNoOfRecords.Text = "Initializing...   Step2"
            lbNoOfRecords.Refresh()
            Call CutWithinLimit()
        End If

        
        '        Dim ColumnName As String = "Time"
        '        Dim SheetName As String = "bts$"
        '        Dim A_party As String = "a"
        '        Dim intialTime As String = dtp_Initial_Time.Value.ToOADate
        '        Dim endTime As String = dtp_Ending_Time.Value.ToOADate
        '        Dim misValue As Object = System.Reflection.Missing.Value
        '        If chkCNIC.Checked = True Then
        '            Call DropTempTable("tblCNIC")
        '            Call createTblCNIC("tblCNIC")
        '        End If

        '        ' Dim TimeInInteger As Integer = CType(intialTime)
        '        Dim queryString As String = "Select * from [" & SheetName & "] where [" & ColumnName & "] between '" & intialTime & "' AND '" & endTime & "'" 'Order By '" & A_party & "'"
        '        'Dim queryString As String = "Select * from [" & SheetName & "] where [" & ColumnName & "] < '" & intialTime & "'" 'AND [" & ColumnName & "]  > '" & endTime & "'"
        '        ' Dim queryString As String = "select [" & ColumnName & "] as party, count(*) as CountOf from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        '        'Dim CountqueryString As String = "select COUNT(*) as CountOf from (select [" & ColumnName1 & "] as Party from [" & SheetName & "] GROUP BY [" & ColumnName1 & "]" + _
        '        '               "UNION ALL select [" & ColumnName2 & "] as Party from [" & SheetName & "] GROUP BY [" & ColumnName2 & "] )"

        '        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & txt_BTS_Path.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        '        'Dim o As OleDbConnection
        '        ' Dim ConnectionStringCopy As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & MakeCopyOfExcel(txt_BTS_Path.Text) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        '        Dim o As New OleDb.OleDbConnection(ConnectionString)
        '        'Dim CopyConn As New OleDb.OleDbConnection(ConnectionStringCopy)
        '        'CopyConn.Open()
        '        o.Open()
        '        Dim FilterQueryString As String

        '        Dim oreader As OleDb.OleDbDataReader
        '        ''FilterQueryString = "Select Distinct(a) from [" & SheetName & "] where ([" & ColumnName & "] < '" & intialTime & "' OR [" & ColumnName & "] > '" & endTime & "')"
        '        ''Dim Ocmd As New OleDb.OleDbCommand(FilterQueryString, o)
        '        ''oreader = Ocmd.ExecuteReader()



        'Call OutOfLimmitConn()
        '        ' '' '' '' ''Dim InserOutOfLimitCommand As SqlCommand
        '        ' '' '' '' ''Dim SqlOutOfLimitReder As SqlDataReader


        '        ' '' '' '' ''Dim currentPhnNumber As Double

        '        ' '' '' '' ''While oreader.Read
        '        ' '' '' '' ''    If IsDBNull(oreader(A_party)) = False Then
        '        ' '' '' '' ''        currentPhnNumber = oreader(A_party)
        '        ' '' '' '' ''        InsertQuery = "INSERT INTO OutOfLimit VALUES ('" & currentPhnNumber & "')"
        '        ' '' '' '' ''        '" & CNIC & "'"
        '        ' '' '' '' ''        InserOutOfLimitCommand = New SqlCommand(InsertQuery, OutOfLimitConnection)
        '        ' '' '' '' ''        InserOutOfLimitCommand.ExecuteNonQuery()
        '        ' '' '' '' ''    End If
        '        ' '' '' '' ''End While
        '        ' '' '' '' ''oreader.Close()
        '        'TargetConnection.Close()
        '        'TargetConnection.Dispose()
        '        '''''' InserOutOfLimitCommand.Dispose()



        '        'Dim ocmd As New OleDb.OleDbCommand(queryString, o)
        '        Dim ocmd1 As New OleDb.OleDbCommand(queryString, o)
        '        'ocmd1.ExecuteNonQuery()

        '        Dim oreader1 As OleDb.OleDbDataReader
        '        oreader1 = ocmd1.ExecuteReader()
        '        Call CreateNewExcelFile()
        '        Dim ColNumber As Integer = 1
        '        Dim ColName As String
        '        Dim ColRange As String
        '        Dim FieldsNames As String
        '        Dim totalfiels As Integer = oreader1.FieldCount
        '        If chkCNIC.Checked = False Then
        '            For i As Integer = 1 To totalfiels
        '                newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 1).ToString

        '            Next

        '        Else
        '            For i As Integer = 1 To totalfiels + 2
        '                If i = 1 Then
        '                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 1).ToString
        '                    FieldsNames = "[" & oreader1.GetName(i - 1).ToString & "] float"
        '                ElseIf i = 2 Then
        '                    newXlWorkSheet.Cells(1, i) = "CNIC_a"
        '                    FieldsNames = FieldsNames & ", [CNIC_a] varchar(255)"
        '                ElseIf i = 3 Then
        '                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 2).ToString
        '                    FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 2).ToString & "] float"
        '                ElseIf i = 4 Then
        '                    newXlWorkSheet.Cells(1, i) = "CNIC_b"
        '                    FieldsNames = FieldsNames & ", [CNIC_b] varchar(255)"

        '                ElseIf i >= 5 Then
        '                    newXlWorkSheet.Cells(1, i) = oreader1.GetName(i - 3).ToString
        '                    FieldsNames = FieldsNames & ", [" & oreader1.GetName(i - 3).ToString & "] varchar(255)"
        '                End If
        '                'If oreader1.GetName(i - 1).ToString = "a" Then
        '                '    ColName = Chr(65 + ColNumber - 1)
        '                '    ColRange = ColName & ":" & ColName
        '                '    newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '                'ElseIf oreader1.GetName(i - 1).ToString = "Time" Then
        '                '    ColName = Chr(65 + ColNumber - 1)
        '                '    ColRange = ColName & ":" & ColName
        '                '    newXlWorkSheet.Range(ColRange).NumberFormat = "h:mm:ss AM/PM"
        '                'ElseIf oreader1.GetName(i - 1).ToString = "b" Then
        '                '    ColName = Chr(65 + ColNumber - 1)
        '                '    ColRange = ColName & ":" & ColName
        '                '    newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '                'End If
        '                'ColNumber = ColNumber + 1
        '            Next
        '            '''' FieldsNames = FieldsNames & ")"
        '            Call CreateTblLimimtedActivity("TblWithinLimit", FieldsNames)
        '        End If
        '        Dim PrevousNumber As String = Nothing
        '        Dim currentNumber As Double
        '        Dim fieldName As String = oreader1.GetName(0).ToString
        '        Dim IndexOfRow As Integer = 2
        '        Dim IsCulprit As Boolean = False
        '        While oreader1.Read
        '            'If PrevousNumber = oreader1(A_party) Then
        '            '    If IsCulprit = False Then
        '            '        GoTo NextRecord
        '            '    ElseIf IsCulprit = True Then
        '            '        GoTo addRecord

        '            '    End If

        '            'End If
        '            ''Application.DoEvents()
        '            ''If StopAnalyze = True Then
        '            ''    Button1.Text = "Close"
        '            ''    Exit While
        '            ''End If
        '            '            IsCulprit = False
        '            If IsDBNull(oreader1(A_party)) = False Then
        '                currentNumber = oreader1(A_party)

        '                'Old purification method

        '                FilterQueryString = "Select * from [" & SheetName & "] where ([" & ColumnName & "] < '" & intialTime & "' OR [" & ColumnName & "] > '" & endTime & "') AND [" & A_party & "] = " & currentNumber & ""
        '                Dim Ocmd As New OleDb.OleDbCommand(FilterQueryString, o)
        '                oreader = Ocmd.ExecuteReader()
        '                CheckedTotal = CheckedTotal + 1
        '                lbBT_CRP_Checked.Text = CheckedTotal
        '                lbBT_CRP_Checked.Refresh()
        '                While oreader.Read
        '                    GoTo NextRecord
        '                End While
        '                oreader.Close()

        '                'New purification method

        '                '' '' '' ''FilterQueryString = "Select a from OutOfLimit where a=" & currentNumber & ""
        '                '' '' '' ''InserOutOfLimitCommand = New SqlCommand(FilterQueryString, OutOfLimitConnection)
        '                '' '' '' ''SqlOutOfLimitReder = InserOutOfLimitCommand.ExecuteReader
        '                '' '' '' ''CheckedTotal = CheckedTotal + 1
        '                '' '' '' ''lbBT_CRP_Checked.Text = CheckedTotal
        '                '' '' '' ''lbBT_CRP_Checked.Refresh()
        '                '' '' '' ''While SqlOutOfLimitReder.Read
        '                '' '' '' ''    GoTo NextRecord
        '                '' '' '' ''End While

        'addRecord:
        '                If chkCNIC.Checked = False Then
        '                    For j As Integer = 1 To totalfiels
        '                        newXlWorkSheet.Cells(IndexOfRow, j) = oreader1(j - 1)
        '                        IsCulprit = True
        '                    Next
        '                    IndexOfRow = IndexOfRow + 1
        '                    lbNoOfRecords.Text = IndexOfRow - 2
        '                Else
        '                    GprCNIC = Nothing
        '                    For j As Integer = 1 To totalfiels + 2
        '                        If j = 1 Then
        '                            newXlWorkSheet.Cells(IndexOfRow, j) = oreader1(j - 1)
        '                            If IsDBNull(oreader1(j - 1)) = False Then
        '                                IsCNIC_Need = True
        '                                Call GetRptNumberCNIC(oreader1(j - 1).ToString)
        '                                If GprCNIC = Nothing Then
        '                                    Call FindPhoneNumber(oreader1(j - 1))
        '                                    If GprCNIC = Nothing Then
        '                                        GprCNIC = "NA"
        '                                    End If
        '                                    Call InsertCNIC_CropBTS(oreader1(j - 1).ToString, GprCNIC.ToString, "tblCNIC")
        '                                End If
        '                                IsCNIC_Need = False
        '                            End If
        '                        ElseIf j = 2 Then
        '                            If GprCNIC = Nothing Then
        '                                GprCNIC = "NA"
        '                            End If
        '                            newXlWorkSheet.Cells(IndexOfRow, j) = GprCNIC.ToString

        '                            GprCNIC = Nothing
        '                        ElseIf j = 3 Then
        '                            newXlWorkSheet.Cells(IndexOfRow, j) = oreader1(j - 2)
        '                            If IsDBNull(oreader1(j - 2)) = False Then
        '                                IsCNIC_Need = True
        '                                Call GetRptNumberCNIC(oreader1(j - 2).ToString)
        '                                If GprCNIC = Nothing Then
        '                                    Call FindPhoneNumber(oreader1(j - 2))
        '                                    If GprCNIC = Nothing Then
        '                                        GprCNIC = "NA"
        '                                    End If
        '                                    Call InsertCNIC_CropBTS(oreader1(j - 2).ToString, GprCNIC.ToString, "tblCNIC")
        '                                End If
        '                                IsCNIC_Need = False
        '                            End If
        '                        ElseIf j = 4 Then
        '                            If GprCNIC = Nothing Then
        '                                GprCNIC = "NA"
        '                            End If
        '                            newXlWorkSheet.Cells(IndexOfRow, j) = GprCNIC.ToString
        '                            GprCNIC = Nothing
        '                        ElseIf j >= 5 Then
        '                            newXlWorkSheet.Cells(IndexOfRow, j) = oreader1(j - 3)
        '                        End If

        '                        IsCulprit = True
        '                    Next
        '                    IndexOfRow = IndexOfRow + 1
        '                    lbNoOfRecords.Text = IndexOfRow - 2
        '                End If
        'NextRecord:
        '                oreader.Close()
        '                'SqlOutOfLimitReder.Close()
        '            End If
        '            'PrevousNumber = oreader1(A_party).ToString
        '        End While
        '        ''TargetConnection.Close()
        '        ''TargetConnection.Dispose()
        '        For i As Integer = 1 To totalfiels
        '            'newXlWorkSheet.Cells(1, i).value()
        '            If newXlWorkSheet.Cells(1, i).value = "a" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '            ElseIf newXlWorkSheet.Cells(1, i).value = "CNIC_a" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '            ElseIf newXlWorkSheet.Cells(1, i).value = "CNIC_b" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '            ElseIf newXlWorkSheet.Cells(1, i).value = "Time" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "hh:mm:ss AM/PM"
        '            ElseIf newXlWorkSheet.Cells(1, i).value = "Date" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "dd-mm-yy"
        '            ElseIf newXlWorkSheet.Cells(1, i).value = "b" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '            ElseIf newXlWorkSheet.Cells(1, i).value = "IMEI" Then
        '                ColName = Chr(65 + ColNumber - 1)
        '                ColRange = ColName & ":" & ColName
        '                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
        '            End If


        '            ColNumber = ColNumber + 1
        '        Next
        '        Try
        '            newXlWorkSheet.Columns.AutoFit()
        '        Catch ex As Exception

        '        End Try
        '        Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(txt_BTS_Path.Text)
        '        Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(txt_BTS_Path.Text)
        '        Dim JustExtention As String = System.IO.Path.GetExtension(txt_BTS_Path.Text)
        '        newXlApp.DisplayAlerts = False
        '        Try
        '            If chkCNIC.Checked = False Then
        '                newXlWorkSheet.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention)
        '            Else
        '                newXlWorkSheet.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)CNIC" & JustExtention)
        '            End If
        '        Catch ex As Exception
        '            If chkCNIC.Checked = False Then
        '                MsgBox("Please close the file" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.Information)
        '            Else
        '                MsgBox("Please close the file" & JustFileName & "(Analyzed)CNIC" & JustExtention, MsgBoxStyle.Information)
        '            End If
        '        End Try

        '        'newXlWorkbook.SaveAs(DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        '        newXlWorkbook.Close()
        '        newXlWorkbooks = Nothing
        '        newXlApp.Quit()
        '        newXlApp = Nothing
        '        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlApp) : newXlApp = Nothing
        '        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkbook) : newXlWorkbook = Nothing
        '        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkbooks) : newXlWorkbooks = Nothing
        '        'System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlWorkSheet) : newXlWorkSheet = Nothing
        '        ' MsgBox("okay" & NumberOfRows, MsgBoxStyle.OkOnly)
        '        oreader1.Close()
        '        ' TargetConnection.Close()
        '        'TargetConnection.Dispose()
        '        OutOfLimitConnection.Close()
        '        OutOfLimitConnection.Dispose()
        '        ' oreader.Close()
        '        ocmd1.Dispose()
        '        If o.State <> ConnectionState.Closed Then
        '            o.Close()
        '            o.Dispose()

        '        End If

        '        'o = Nothing
        '        Call ChangeTimeFormat("h:mm:ss AM/PM")
        '        Me.Cursor = Cursors.Default
        '        Button1.Text = "Close"
        '        ' MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.OkOnly)
        '        If chkCNIC.Checked = False Then
        '            lbfilepath.Text = DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention
        '        Else
        '            lbfilepath.Text = DirectoryPath & "\" & JustFileName & "(Analyzed)CNIC" & JustExtention
        '        End If
        '        If cbFind_b_in_a.Checked = True Then
        '            ' Call Find_b_in_a(lbfilepath.Text)
        '            NumberOfFiles = 1
        '            AllCommonFileNames(0) = lbfilepath.Text
        '            Call FindGroupsOfOne()
        '            MsgBox("Both files 'Analyzed and Analyzed Groups' have been created", MsgBoxStyle.Information)
        '            ' MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.OkOnly)
        '        Else
        '            MsgBox("File has been Created:" & vbCrLf & DirectoryPath & "\" & JustFileName & "(Analyzed)" & JustExtention, MsgBoxStyle.Information)
        '        End If
    End Sub
    Public newXlApp As Excel.Application
    Public newXlWorkbooks As Excel.Workbooks
    Public newXlWorkbook As Excel.Workbook
    Public newXlWorkSheet As Excel.Worksheet
    Public xlSheets As Excel.Worksheets
    Sub CreateNewExcelFile()
        newXlApp = New Excel.Application
        newXlWorkbook = newXlApp.Workbooks.Add()

        'newXlWorkSheet = "bts"
        newXlWorkSheet = newXlWorkbook.Sheets("Sheet1")
        newXlWorkSheet.Name = "bts"
    End Sub
    Sub FindGroups()
        Call CreateTblDuplicateCheck(TblCheckDuplication)
        If chkLimitedActivity.Checked = False Then
            Dim Reply As MsgBoxResult
            Reply = MsgBox("Do you want to create file 'Limited Activity'", MsgBoxStyle.YesNoCancel)
            If Reply = MsgBoxResult.Yes Then
                chkLimitedActivity.Checked = True
            ElseIf Reply = MsgBoxResult.No Then
                chkLimitedActivity.Checked = False
            ElseIf Reply = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If


        lbTotalNos.Text = "Initializing..."
        lbCheckedNos.Text = "Initializing..."
        lbFound.Text = "0"
        Dim OnlyFileNames(NumberOfFiles) As String
        Dim chkCNIC_Exist(NumberOfFiles) As Boolean
        Dim ConnectionString As String
        Dim PhoneNumber As String
        Dim SheetName1 As String = "bts$"
        Dim SheetName2 As String = "bts$"
        Dim SheetName As String = "bts$"
        Dim ocmd_a As OleDbCommand
        Dim ocmd_b As OleDbCommand
        Dim ocmdDumi As OleDbCommand
        Dim DumiReader As OleDbDataReader
        Dim NumberOfFields As Integer = 0
        Dim FileIndex As Integer = 0
        Dim FileName As String = ""
        Dim isErrorInFile As Boolean = False
        Dim SaveFilePath As String = FolderPath & "\Multi BTS Groups" & SaveGroupExtention
        'Dim xlApp As New Excel.Application
        'Dim xlWb As Excel.Workbook
        'xlApp = GetObject(, "Excel.Application")
        'For Each xlWb In xlApp.Workbooks
        '    xlWb.Close(False)
        'Next

        'xlApp = Nothing
        'Try
        '    xlApp.Workbooks.Close()
        '    'xlWb.Close(False)
        'Catch ex As Exception
        '    MsgBox("evrror", MsgBoxStyle.Information)
        'End Try
        'xlApp = Nothing
        'xlWb = Nothing


        'newXlWorkbook.Close(False, SaveFilePath)
        Dim GroupCon As OleDbConnection
        'xlWorkbooks = xlApp.Workbooks
        Dim IsFileWithCNIC As Boolean = False
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Dim queryString_aDumi As String
        'xlApp.DisplayAlerts = False
        Dim fileNameAndPath As String = AllCommonFileNames(0)

        For j As Integer = 0 To NumberOfFiles - 1
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
            Dim o1 As New OleDb.OleDbConnection(ConnectionString)
            o1.Open()
            'Dim Indexofdot As Integer = AllCommonFileNames(j).IndexOf(".")
            'If AllCommonFileNames(j).Substring(AllCommonFileNames(j).IndexOf(".") - 4, 4) = "CNIC" Then
            If AllCommonFileNames(j).Contains("CNIC.xlsx") Or AllCommonFileNames(j).Contains("CNIC.xls") Then
                queryString_aDumi = "Select a,CNIC_a,b,CNIC_b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "]"
                IsFileWithCNIC = True
                chkCNIC_Exist(j) = True
            Else
                queryString_aDumi = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "]"
                chkCNIC_Exist(j) = False
            End If
            'queryString_aDumi = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "]"
            ocmdDumi = New OleDb.OleDbCommand(queryString_aDumi, o1)
            Try
                DumiReader = ocmdDumi.ExecuteReader()
            Catch ex As Exception
                isErrorInFile = True
                'FileIndex += 1
                'ReDim Preserve FileName(FileIndex)
                FileName = FileName & System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j)) & ", "
            End Try
            DumiReader.Close()
            ocmdDumi.Dispose()
            o1.Close()
        Next
        If isErrorInFile = True Then
            MsgBox("Please make sure of setting of Files " & FileName & vbCrLf & "Sheet Name: bts" & vbCrLf & "Column Names: a,b,Time,Date,Call Type,Duration,Cell ID,IMEI,IMSI,Site", MsgBoxStyle.Information)
            Exit Sub
        End If
        'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
        'Dim o1 As New OleDb.OleDbConnection(ConnectionString)
        'o1.Open()

        fileNameAndPath = fileNameAndPath.Insert(fileNameAndPath.IndexOf(".") - 1, " Groups")
        ' xlWorkbook = xlApp.Workbooks.Open(lbfilepath.Text, , , , , , , , , True, False)
        Call CreateNewExcelFile()
        ' If NumberOfFiles > 1 Then
        If IsFileWithCNIC = False Then
            newXlWorkSheet.Cells(1, 1) = "a"
            newXlWorkSheet.Range("A:A").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 2) = "b"
            newXlWorkSheet.Range("B:B").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 3) = "Time"
            newXlWorkSheet.Range("C:C").NumberFormat = "hh:mm:ss AM/PM"
            newXlWorkSheet.Cells(1, 4) = "Date"
            newXlWorkSheet.Range("D:D").NumberFormat = "dd-mm-yy"
            newXlWorkSheet.Cells(1, 5) = "Call Type"
            newXlWorkSheet.Cells(1, 6) = "Duration"
            newXlWorkSheet.Cells(1, 7) = "Cell ID"
            newXlWorkSheet.Cells(1, 8) = "IMEI"
            newXlWorkSheet.Range("H:H").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 9) = "IMSI"
            newXlWorkSheet.Range("I:I").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 10) = "Site"
            NumberOfFields = 10
        Else
            newXlWorkSheet.Cells(1, 1) = "a"
            newXlWorkSheet.Range("A:A").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 2) = "CNIC_a"
            newXlWorkSheet.Range("B:B").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 3) = "b"
            newXlWorkSheet.Range("C:C").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 4) = "CNIC_b"
            newXlWorkSheet.Range("D:D").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 5) = "Time"
            newXlWorkSheet.Range("E:E").NumberFormat = "hh:mm:ss AM/PM"
            newXlWorkSheet.Cells(1, 6) = "Date"
            newXlWorkSheet.Range("F:F").NumberFormat = "dd-mm-yy"
            newXlWorkSheet.Cells(1, 7) = "Call Type"
            newXlWorkSheet.Cells(1, 8) = "Duration"
            newXlWorkSheet.Cells(1, 9) = "Cell ID"
            newXlWorkSheet.Cells(1, 10) = "IMEI"
            newXlWorkSheet.Range("J:J").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 11) = "IMSI"
            newXlWorkSheet.Range("K:K").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 12) = "Site"
            NumberOfFields = 12
        End If
        'Else
        'Dim queryString_bOne As String = "Select * from [" & SheetName1 & "]"

        'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
        'Dim o As New OleDb.OleDbConnection(ConnectionString)
        'o.Open()

        'Dim ocmd1 As New OleDb.OleDbCommand(queryString_bOne, o)

        'Dim oreader_bone As OleDb.OleDbDataReader
        'oreader_bone = ocmd1.ExecuteReader()
        'NumberOfFields = oreader_bone.FieldCount
        'For j As Integer = 1 To NumberOfFields
        '    newXlWorkSheet.Cells(1, j) = oreader_bone.GetName(j - 1).ToString
        'Next
        'oreader_bone.Close()
        'ocmd1.Dispose()
        'o.Close()
        'o.Dispose()
        'End If
        Call DropTempTable("bParty")
        Call Create_bParyt_Table("bParty")
        Dim ConnString(NumberOfFiles) As String
        Dim GroupConnection(NumberOfFiles) As OleDb.OleDbConnection
        Dim GroupCommand(NumberOfFiles) As OleDb.OleDbCommand
        Dim GroupReader(NumberOfFiles) As OleDb.OleDbDataReader

        For j As Integer = 0 To NumberOfFiles - 1
            ' OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
            ConnString(j) = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
            GroupConnection(j) = New OleDb.OleDbConnection(ConnString(j))
            GroupConnection(j).Open()
        Next

        For j As Integer = 0 To NumberOfFiles - 1
            OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
            'CommonNoTable.Rows(1).Cells(j + 3).Range.Text = OnlyFileNames(j)
            'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
            'Dim o1 As New OleDb.OleDbConnection(ConnectionString)
            'o1.Open()
            QueryString = "Select DISTINCT(b) from [" & SheetName & "]"
            'Dim InsertCommand As New OleDb.OleDbCommand(QueryString, o1)
            'Dim InsertReader As OleDb.OleDbDataReader
            GroupCommand(j) = New OleDb.OleDbCommand(QueryString, GroupConnection(j))

            Try
                ' InsertReader = InsertCommand.ExecuteReader()
                GroupReader(j) = GroupCommand(j).ExecuteReader()
            Catch ex As Exception
                Call DropTempTable("bParty")
                Me.Cursor = Cursors.Default
                'MsgBox("The result has not been produced" & vbCrLf & "Please make sure sheet is 'Sheet1' with columns name 'a' , 'b' of the file " & TempTableName.Substring(1), MsgBoxStyle.Information)
                Exit Sub
            End Try
            While GroupReader(j).Read
                If IsDBNull(GroupReader(j)(0)) = False Then
                    PhoneNumber = GroupReader(j)(0)
                    If PhoneNumber.Length >= 10 Then
                        Call InsertPhoneNumber(PhoneNumber.Substring(PhoneNumber.Length - 10, 10))
                    End If
                End If
            End While
            GroupReader(j).Close()
            GroupCommand(j).Dispose()
            'o1.Close()
            'o1.Dispose()
            'While InsertReader.Read
            '    If IsDBNull(InsertReader(0)) = False Then
            '        PhoneNumber = InsertReader(0)
            '        If PhoneNumber.Length >= 10 Then
            '            Call InsertPhoneNumber(PhoneNumber.Substring(PhoneNumber.Length - 10, 10))
            '        End If
            '    End If
            'End While
            'InsertReader.Close()
            'o1.Close()
            'o1.Dispose()

        Next


        newXlApp.DisplayAlerts = False
        Dim CommonNoQuery As String
        Dim TotalNumbers As Integer = 0
        Dim queryString_b As String = "Select DISTINCT(bParty) from bParty order by bParty"
        Dim bPartyOthersCommand As SqlCommand
        Dim bPartyOthersReader As SqlDataReader
        Dim queryString_a As String = Nothing
        Dim queryString_Full_b As String = Nothing
        Dim oreader_a As OleDb.OleDbDataReader
        Dim oreader_b As OleDb.OleDbDataReader
        Dim oreader_full_b As OleDb.OleDbDataReader
        Dim checkedNumber As Integer = 0
        Dim totalfields As Integer = 0
        Dim IndexOfRow As Integer = 2
        Dim PreNumber As Long = 0
        Dim length As Integer = 0
        Dim IsSetFieldsName As Boolean = False
        Call ConnectionOpen()
        CommonNoQuery = "Select DISTINCT(bParty) from bParty order by bParty"
        bPartyOthersCommand = New SqlCommand(CommonNoQuery, TargetConnection)
        bPartyOthersReader = bPartyOthersCommand.ExecuteReader()
        While bPartyOthersReader.Read()
            TotalNumbers += 1
            lbTotalNos.Text = TotalNumbers
            lbTotalNos.Refresh()
        End While

        bPartyOthersReader.Close()

        bPartyOthersCommand = New SqlCommand(CommonNoQuery, TargetConnection)
        bPartyOthersReader = bPartyOthersCommand.ExecuteReader()
        Dim IsbpartyAdded As Boolean = False
        Dim NextNumber As Long = Nothing
        Dim NextNumbertxt As String = Nothing
        Dim isNumberFound As Integer = False
        Dim a As String = Nothing
        Dim b As String = Nothing
        Dim grpTime As String = Nothing
        Dim grpDate As String = Nothing
        Dim Duration As String = Nothing
        While bPartyOthersReader.Read()
            If IsDBNull(bPartyOthersReader(0)) = False Then
                isNumberFound = False
                For k As Integer = 0 To NumberOfFiles - 1
                    'OnlyFileNames(k) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(k))
                    'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(k) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
                    'GroupCon = New OleDb.OleDbConnection(ConnectionString)
                    'GroupCon.Open()
                    'If NumberOfFiles = 1 Then
                    '    queryString_a = "Select * from [" & SheetName1 & "] where a like '%" & bPartyOthersReader(0) & "'"
                    'Else
                    '    queryString_a = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "] where a like '%" & bPartyOthersReader(0) & "'"
                    'End If
                    If chkCNIC_Exist(k) = False Then
                        queryString_a = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] where a like '%" & bPartyOthersReader(0) & "'"
                    Else
                        'ocmd_a = New OleDb.OleDbCommand(queryString_a, GroupCon)
                        queryString_a = "Select a,CNIC_a,b,CNIC_b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] where a like '%" & bPartyOthersReader(0) & "'"
                    End If
                    'oreader_a = ocmd_a.ExecuteReader()
                    ' isNumberFound = False

                    GroupCommand(k) = New OleDb.OleDbCommand(queryString_a, GroupConnection(k))


                    ' InsertReader = InsertCommand.ExecuteReader()
                    GroupReader(k) = GroupCommand(k).ExecuteReader()
                    While GroupReader(k).Read
                        If IsDBNull(GroupReader(k)(0)) = False Then
                            isNumberFound = True
                            If IsDBNull(GroupReader(k)("a")) = False Then
                                a = GroupReader(k)("a")
                            Else
                                a = "NULL"
                            End If
                            If IsDBNull(GroupReader(k)("b")) = False Then
                                b = GroupReader(k)("b")
                            Else
                                b = "NULL"
                            End If
                            If IsDBNull(GroupReader(k)("Time")) = False Then
                                grpTime = GroupReader(k)("Time")
                            Else
                                grpTime = "NULL"
                            End If
                            If IsDBNull(GroupReader(k)("Date")) = False Then
                                grpDate = GroupReader(k)("Date")
                            Else
                                grpDate = "NULL"
                            End If
                            If IsDBNull(GroupReader(k)("Duration")) = False Then
                                Duration = GroupReader(k)("Duration")
                            Else
                                Duration = "NULL"
                            End If

                            If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                                If IsFileWithCNIC = False Then
                                    For j As Integer = 1 To NumberOfFields
                                        newXlWorkSheet.Cells(IndexOfRow, j) = GroupReader(k)(j - 1)
                                    Next
                                Else
                                    If chkCNIC_Exist(k) = False Then
                                        For j As Integer = 1 To NumberOfFields - 2

                                            If j = 1 Then
                                                newXlWorkSheet.Cells(IndexOfRow, j) = GroupReader(k)(j - 1)
                                            ElseIf j = 2 Then
                                                newXlWorkSheet.Cells(IndexOfRow, j + 1) = GroupReader(k)(j - 1)
                                            ElseIf j >= 3 Then
                                                newXlWorkSheet.Cells(IndexOfRow, j + 2) = GroupReader(k)(j - 1)
                                            End If
                                        Next
                                    Else
                                        For j As Integer = 1 To NumberOfFields
                                            newXlWorkSheet.Cells(IndexOfRow, j) = GroupReader(k)(j - 1)
                                        Next
                                    End If
                                End If
                                lbFound.Text = IndexOfRow - 1
                                IndexOfRow = IndexOfRow + 1
                            End If
                        End If
                    End While
                    GroupReader(k).Close()
                    GroupCommand(k).Dispose()
                    'oreader_a.Close()
                    'oreader_b.Close()
                    'ocmd_a.Dispose()
                    'ocmd_b.Dispose()
                    'GroupCon.Close()
                    'GroupCon.Dispose()

                Next
                If isNumberFound = True Then
                    For m As Integer = 0 To NumberOfFiles - 1
                        'OnlyFileNames(m) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(m))
                        'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(m) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
                        'GroupCon = New OleDb.OleDbConnection(ConnectionString)
                        'GroupCon.Open()
                        'If NumberOfFiles = 1 Then
                        '    queryString_b = "Select * from [" & SheetName1 & "] Where b like '%" & bPartyOthersReader(0) & "'"
                        'Else
                        '    queryString_b = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "] Where b like '%" & bPartyOthersReader(0) & "'"
                        'End If
                        If chkCNIC_Exist(m) = False Then
                            queryString_b = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] Where b like '%" & bPartyOthersReader(0) & "'"
                        Else
                            queryString_b = "Select a,CNIC_a,b,CNIC_b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] Where b like '%" & bPartyOthersReader(0) & "'"
                        End If

                        'ocmd_b = New OleDb.OleDbCommand(queryString_b, GroupCon)
                        'oreader_b = ocmd_b.ExecuteReader()
                        GroupCommand(m) = New OleDb.OleDbCommand(queryString_b, GroupConnection(m))


                        ' InsertReader = InsertCommand.ExecuteReader()
                        GroupReader(m) = GroupCommand(m).ExecuteReader()
                        While GroupReader(m).Read
                            If IsDBNull(GroupReader(m)("a")) = False Then
                                a = GroupReader(m)("a")
                            Else
                                a = "NULL"
                            End If
                            If IsDBNull(GroupReader(m)("b")) = False Then
                                b = GroupReader(m)("b")
                            Else
                                b = "NULL"
                            End If
                            If IsDBNull(GroupReader(m)("Time")) = False Then
                                grpTime = GroupReader(m)("Time")
                            Else
                                grpTime = "NULL"
                            End If
                            If IsDBNull(GroupReader(m)("Date")) = False Then
                                grpDate = GroupReader(m)("Date")
                            Else
                                grpDate = "NULL"
                            End If
                            If IsDBNull(GroupReader(m)("Duration")) = False Then
                                Duration = GroupReader(m)("Duration")
                            Else
                                Duration = "NULL"
                            End If

                            If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                                If IsFileWithCNIC = False Then
                                    For j As Integer = 1 To NumberOfFields
                                        newXlWorkSheet.Cells(IndexOfRow, j) = GroupReader(m)(j - 1)
                                    Next
                                Else
                                    If chkCNIC_Exist(m) = False Then
                                        For j As Integer = 1 To NumberOfFields - 2

                                            If j = 1 Then
                                                newXlWorkSheet.Cells(IndexOfRow, j) = GroupReader(m)(j - 1)
                                            ElseIf j = 2 Then
                                                newXlWorkSheet.Cells(IndexOfRow, j + 1) = GroupReader(m)(j - 1)
                                            ElseIf j >= 3 Then
                                                newXlWorkSheet.Cells(IndexOfRow, j + 2) = GroupReader(m)(j - 1)
                                            End If
                                        Next
                                    Else
                                        For j As Integer = 1 To NumberOfFields
                                            newXlWorkSheet.Cells(IndexOfRow, j) = GroupReader(m)(j - 1)
                                        Next
                                    End If
                                End If
                                lbFound.Text = IndexOfRow - 1
                                lbFound.Refresh()
                                IndexOfRow = IndexOfRow + 1
                            End If
                        End While
                        GroupReader(m).Close()
                        GroupCommand(m).Dispose()
                        'oreader_a.Close()
                        'oreader_b.Close()
                        ''ocmd_a.Dispose()
                        'ocmd_b.Dispose()
                        'GroupCon.Close()
                        'GroupCon.Dispose()
                    Next

                End If

            End If
            checkedNumber += 1
            lbCheckedNos.Text = checkedNumber
            lbCheckedNos.Refresh()
        End While
        TargetConnection.Close()
        TargetConnection.Dispose()
        newXlApp.DisplayAlerts = False
        If chkLimitedActivity.Checked = False Then
            ' fileNameAndPath = fileNameAndPath.Insert(fileNameAndPath.IndexOf(".") - 1, " Groups")
            For j As Integer = 0 To NumberOfFiles - 1
                ' OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
                GroupConnection(j).Close()

            Next
            Try
                'If NumberOfFiles = 1 Then
                '    Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(AllCommonFileNames(0))
                '    Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(0))
                '    Dim JustExtention As String = System.IO.Path.GetExtension(AllCommonFileNames(0))
                '    newXlWorkSheet.SaveAs(AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups"))
                '    MsgBox("File has been created: " & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups"), MsgBoxStyle.Information)
                'Else
                newXlWorkSheet.Columns.AutoFit()
                newXlWorkSheet.SaveAs(SaveFilePath)

                MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
                ' End If
                newXlWorkbook.Close()
                newXlWorkbooks = Nothing
                newXlApp.Quit()
                newXlApp = Nothing
                Exit Sub
                'MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Please close the file" & SaveFilePath, MsgBoxStyle.Information)
                newXlWorkbook.Close()
                newXlWorkbooks = Nothing
                newXlApp.Quit()
                newXlApp = Nothing
                Exit Sub
            End Try
        Else
            Try
                newXlWorkSheet.Columns.AutoFit()
                newXlWorkSheet.SaveAs(SaveFilePath)
                newXlWorkbook.Close()
                newXlWorkbooks = Nothing
                newXlApp.Quit()
                newXlApp = Nothing
            Catch ex As Exception
                MsgBox("Please close the file" & SaveFilePath, MsgBoxStyle.Information)
                newXlWorkbook.Close()
                newXlWorkbooks = Nothing
                newXlApp.Quit()
                newXlApp = Nothing
                Exit Sub
            End Try
        End If
        'newXlWorkbook.Close()
        'newXlWorkbooks = Nothing
        'newXlApp.Quit()
        'newXlApp = Nothing
        lbTotalNos.Text = "Initializing..."
        lbCheckedNos.Text = "Initializing..."
        lbFound.Text = "0"
        Call CreateNewExcelFile()
        If IsFileWithCNIC = False Then
            newXlWorkSheet.Cells(1, 1) = "a"
            newXlWorkSheet.Range("A:A").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 2) = "b"
            newXlWorkSheet.Range("B:B").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 3) = "Time"
            newXlWorkSheet.Range("C:C").NumberFormat = "hh:mm:ss AM/PM"
            newXlWorkSheet.Cells(1, 4) = "Date"
            newXlWorkSheet.Range("D:D").NumberFormat = "dd-mm-yy"
            newXlWorkSheet.Cells(1, 5) = "Call Type"
            newXlWorkSheet.Cells(1, 6) = "Duration"
            newXlWorkSheet.Cells(1, 7) = "Cell ID"
            newXlWorkSheet.Cells(1, 8) = "IMEI"
            newXlWorkSheet.Range("H:H").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 9) = "IMSI"
            newXlWorkSheet.Range("I:I").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 10) = "Site"
            NumberOfFields = 10
        Else
            newXlWorkSheet.Cells(1, 1) = "a"
            newXlWorkSheet.Range("A:A").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 2) = "CNIC_a"
            newXlWorkSheet.Range("B:B").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 3) = "b"
            newXlWorkSheet.Range("C:C").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 4) = "CNIC_b"
            newXlWorkSheet.Range("D:D").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 5) = "Time"
            newXlWorkSheet.Range("E:E").NumberFormat = "hh:mm:ss AM/PM"
            newXlWorkSheet.Cells(1, 6) = "Date"
            newXlWorkSheet.Range("F:F").NumberFormat = "dd-mm-yy"
            newXlWorkSheet.Cells(1, 7) = "Call Type"
            newXlWorkSheet.Cells(1, 8) = "Duration"
            newXlWorkSheet.Cells(1, 9) = "Cell ID"
            newXlWorkSheet.Cells(1, 10) = "IMEI"
            newXlWorkSheet.Range("J:J").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 11) = "IMSI"
            newXlWorkSheet.Range("K:K").NumberFormat = "0"
            newXlWorkSheet.Cells(1, 12) = "Site"
            NumberOfFields = 12
        End If
        Dim ColumnsHeaders As String = Nothing
        If IsFileWithCNIC = False Then
            ColumnsHeaders = "[a] nvarchar(255),[b] nvarchar(255),[Time] nvarchar(255),[Date] nvarchar(255),[Call Type] nvarchar(255),[Duration] nvarchar(255),[Cell ID] nvarchar(255),[IMEI] nvarchar(255),[IMSI] nvarchar(255),[Site] nvarchar(255)"
            NumberOfFields = 10
        Else
            ColumnsHeaders = "[a] nvarchar(255),[CNIC_a] nvarchar(255),[b] nvarchar(255),[CNIC_b] nvarchar(255),[Time] nvarchar(255),[Date] nvarchar(255),[Call Type] nvarchar(255),[Duration] nvarchar(255),[Cell ID] nvarchar(255),[IMEI] nvarchar(255),[IMSI] nvarchar(255),[Site] nvarchar(255)"
            NumberOfFields = 12
        End If
        Call CreateTblLimimtedActivity("tblLimitedActivity", ColumnsHeaders)
        For k As Integer = 0 To NumberOfFiles - 1
            If IsFileWithCNIC = False Then
                queryString_a = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] " '"
                GroupCommand(k) = New OleDb.OleDbCommand(queryString_a, GroupConnection(k))
                GroupReader(k) = GroupCommand(k).ExecuteReader()
                While GroupReader(k).Read
                    ColumnsHeaders = Nothing
                    For j As Integer = 1 To NumberOfFields
                        If j > 1 Then
                            If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                ColumnsHeaders = ColumnsHeaders & ", '" & GroupReader(k)(j - 1).ToString & "'"
                            Else
                                ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                            End If
                        ElseIf j = 1 Then
                            If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                ColumnsHeaders = "'" & GroupReader(k)(j - 1).ToString & "'"
                            Else
                                ColumnsHeaders = "'NA'"
                            End If
                        End If
                    Next
                    Call InsertAnalyzedToSQL("tblLimitedActivity", ColumnsHeaders)
                    TotalNumbers += 1
                    lbTotalNos.Text = TotalNumbers
                    lbTotalNos.Refresh()
                End While
                GroupReader(k).Close()
                GroupCommand(k).Dispose()
            Else
                If chkCNIC_Exist(k) = False Then
                    queryString_a = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] " '"
                    GroupCommand(k) = New OleDb.OleDbCommand(queryString_a, GroupConnection(k))
                    GroupReader(k) = GroupCommand(k).ExecuteReader()
                    While GroupReader(k).Read
                        ColumnsHeaders = Nothing
                        For j As Integer = 1 To NumberOfFields - 2
                            If j > 2 Then
                                If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                    ColumnsHeaders = ColumnsHeaders & ", '" & GroupReader(k)(j - 1).ToString & "'"
                                Else
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                End If
                            ElseIf j = 1 Then
                                If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                    ColumnsHeaders = "'" & GroupReader(k)(j - 1).ToString & "'"
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                Else
                                    ColumnsHeaders = "'NA'"
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                End If

                            ElseIf j = 2 Then
                                If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                    ColumnsHeaders = ColumnsHeaders & ", '" & GroupReader(k)(j - 1).ToString & "'"
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                Else
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                End If
                            End If
                        Next
                        Call InsertAnalyzedToSQL("tblLimitedActivity", ColumnsHeaders)
                        TotalNumbers += 1
                        lbTotalNos.Text = TotalNumbers
                        lbTotalNos.Refresh()
                    End While
                    GroupReader(k).Close()
                    GroupCommand(k).Dispose()
                Else
                    queryString_a = "Select a,CNIC_a,b,CNIC_b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [" & SheetName1 & "] " '"
                    GroupCommand(k) = New OleDb.OleDbCommand(queryString_a, GroupConnection(k))
                    GroupReader(k) = GroupCommand(k).ExecuteReader()
                    While GroupReader(k).Read
                        ColumnsHeaders = Nothing
                        For j As Integer = 1 To NumberOfFields
                            If j > 1 Then
                                If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                    ColumnsHeaders = ColumnsHeaders & ", '" & GroupReader(k)(j - 1).ToString & "'"
                                Else
                                    ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                                End If
                            ElseIf j = 1 Then
                                If IsDBNull(GroupReader(k)(j - 1)) = False Then
                                    ColumnsHeaders = "'" & GroupReader(k)(j - 1).ToString & "'"
                                Else
                                    ColumnsHeaders = "'NA'"
                                End If
                            End If
                        Next
                        Call InsertAnalyzedToSQL("tblLimitedActivity", ColumnsHeaders)
                        TotalNumbers += 1
                        lbTotalNos.Text = TotalNumbers
                        lbTotalNos.Refresh()
                    End While
                    GroupReader(k).Close()
                    GroupCommand(k).Dispose()
                End If
            End If
        Next
        For j As Integer = 0 To NumberOfFiles - 1
            ' OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
            GroupConnection(j).Close()

        Next
        Dim MultiQueryString As String
        Dim o As OleDbConnection
        Dim ocmd1 As OleDbCommand
        Dim oreader_bone As OleDbDataReader
        MultiQueryString = "Select a,b from [" & SheetName & "]"

        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & SaveFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        o = New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        ocmd1 = New OleDb.OleDbCommand(MultiQueryString, o)
        oreader_bone = ocmd1.ExecuteReader()
        NumberOfFields = oreader_bone.FieldCount
        TotalNumbers = 0
        While oreader_bone.Read
            If IsDBNull(oreader_bone(0)) = False And IsDBNull(oreader_bone(1)) = False Then
                Call RmvGrpFromAnalyzed(oreader_bone(0).ToString, oreader_bone(1).ToString, "tblLimitedActivity")
            ElseIf IsDBNull(oreader_bone(0)) = False And IsDBNull(oreader_bone(1)) = True Then
                Call RmvGrpFromAnalyzed(oreader_bone(0).ToString, "NA", "tblLimitedActivity")
            ElseIf IsDBNull(oreader_bone(0)) = True And IsDBNull(oreader_bone(1)) = False Then
                Call RmvGrpFromAnalyzed("NA", oreader_bone(1).ToString, "tblLimitedActivity")
            End If
            TotalNumbers += 1
            lbCheckedNos.Text = TotalNumbers
            lbCheckedNos.Refresh()
        End While
        oreader_bone.Close()
        ocmd1.Dispose()
        o.Close()
        o.Dispose()
        TotalNumbers = 0
        Call ConnectionOpen()
        QueryString = "Select * from tblLimitedActivity"
        bPartyOthersCommand = New SqlCommand(QueryString, TargetConnection)
        bPartyOthersReader = bPartyOthersCommand.ExecuteReader()
        NumberOfFields = bPartyOthersReader.FieldCount
        IndexOfRow = 2
        While bPartyOthersReader.Read
            For j As Integer = 1 To NumberOfFields
                If bPartyOthersReader(j - 1) <> "NA" Then
                    newXlWorkSheet.Cells(IndexOfRow, j) = bPartyOthersReader(j - 1)
                Else
                    newXlWorkSheet.Cells(IndexOfRow, j) = ""
                End If
            Next
            TotalNumbers += 1
            lbFound.Text = TotalNumbers
            lbFound.Refresh()
            IndexOfRow = IndexOfRow + 1
        End While
        Try
            newXlWorkSheet.Columns.AutoFit()
            newXlApp.DisplayAlerts = False
            Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(AllCommonFileNames(0))
            Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(0))
            Dim JustExtention As String = System.IO.Path.GetExtension(AllCommonFileNames(0))
            newXlWorkSheet.SaveAs(SaveFilePath.Insert(SaveFilePath.IndexOf(".") - 1, " LimitedActivity"))
        Catch ex As Exception
            MsgBox("Please close the file" & SaveFilePath, MsgBoxStyle.Information)
        End Try
        newXlWorkbook.Close()
        newXlWorkbooks = Nothing
        newXlApp.Quit()
        newXlApp = Nothing
        If chkLimitedActivity.Checked = True Then
            MsgBox("File has been created: " & SaveFilePath.Insert(SaveFilePath.IndexOf(".") - 1, " Groups") & vbCrLf & "File has been created: " & SaveFilePath.Insert(SaveFilePath.IndexOf(".") - 1, " LimitedActivity"), MsgBoxStyle.Information)
        End If
    End Sub
    Sub FindGroupsOfOne()
        Call CreateTblDuplicateCheck(TblCheckDuplication)
        If chkLimitedActivity.Checked = False Then
            Dim Reply As MsgBoxResult
            Reply = MsgBox("Do you want to create file 'Limited Activity'", MsgBoxStyle.YesNoCancel)
            If Reply = MsgBoxResult.Yes Then
                chkLimitedActivity.Checked = True
            ElseIf Reply = MsgBoxResult.No Then
                chkLimitedActivity.Checked = False
            ElseIf Reply = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If
        lbTotalNos.Text = "Initializing..."
        lbCheckedNos.Text = "Initializing..."
        lbFound.Text = "0"
        Dim OnlyFileNames(NumberOfFiles) As String
        Dim ConnectionString As String
        Dim PhoneNumber As String
        Dim SheetName1 As String = "Sheet1$"
        Dim SheetName2 As String = "Sheet1$"
        Dim SheetName As String = "bts$"

        Dim ocmd_a As OleDbCommand
        Dim ocmd_b As OleDbCommand
        Dim ocmd_Full_b As OleDbCommand
        Dim NumberOfFields As Integer = 0
        Dim SaveFilePath As String = FolderPath & "\Multi BTS Groups" & SaveGroupExtention
        Dim GroupCon As OleDbConnection
        'xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        'xlApp.DisplayAlerts = False
        Dim fileNameAndPath As String = AllCommonFileNames(0)

        fileNameAndPath = fileNameAndPath.Insert(fileNameAndPath.IndexOf(".") - 1, " Groups")
        ' xlWorkbook = xlApp.Workbooks.Open(lbfilepath.Text, , , , , , , , , True, False)
        Call CreateNewExcelFile()
        Dim NameOfSheet As String = newXlWorkSheet.Name & "$"
        Dim queryString_bOne As String = "Select * from [" & SheetName & "]"

        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()

        Dim ocmd1 As New OleDb.OleDbCommand(queryString_bOne, o)

        Dim oreader_bone As OleDb.OleDbDataReader
        Try
            oreader_bone = ocmd1.ExecuteReader()
        Catch ex As Exception
            Call DropTempTable("bParty")
            Me.Cursor = Cursors.Default
            MsgBox("Please make sure the setting of file: " & System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(0)) & vbCrLf & "Sheet Name: bts" & vbCrLf & "Columns name: a , b", MsgBoxStyle.Information)
            Exit Sub
        End Try
        Dim ColNumber As Integer = 1
        Dim ColName As String = Nothing
        Dim ColRange As String = Nothing
        NumberOfFields = oreader_bone.FieldCount
        For j As Integer = 1 To NumberOfFields
            newXlWorkSheet.Cells(1, j) = oreader_bone.GetName(j - 1).ToString
            If newXlWorkSheet.Cells(1, j).value = "a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "CNIC_a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "CNIC_b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "Time" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "hh:mm:ss AM/PM"
            ElseIf newXlWorkSheet.Cells(1, j).value = "Date" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "dd-mm-yy"
            ElseIf newXlWorkSheet.Cells(1, j).value = "b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "IMEI" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            End If


            ColNumber = ColNumber + 1
        Next
        oreader_bone.Close()
        ' ocmd1.Dispose()
        o.Close()
        ' o.Dispose()

        Call DropTempTable("bParty")
        Call Create_bParyt_Table("bParty")

        ' OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
        'CommonNoTable.Rows(1).Cells(j + 3).Range.Text = OnlyFileNames(j)

        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o1 As New OleDb.OleDbConnection(ConnectionString)
        o1.Open()
        QueryString = "Select DISTINCT(b) from [" & SheetName & "]"
        Dim InsertCommand As New OleDb.OleDbCommand(QueryString, o1)
        Dim InsertReader As OleDb.OleDbDataReader
        Try
            InsertReader = InsertCommand.ExecuteReader()
        Catch ex As Exception
            Call DropTempTable("bParty")
            Me.Cursor = Cursors.Default
            MsgBox("Please make sure the setting of file: " & System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(0)) & vbCrLf & "Sheet Name: bts" & vbCrLf & "Columns name: a , b", MsgBoxStyle.Information)
            Exit Sub
        End Try

        While InsertReader.Read
            If IsDBNull(InsertReader(0)) = False Then
                PhoneNumber = InsertReader(0)
                If PhoneNumber.Length >= 10 Then
                    Call InsertPhoneNumber(PhoneNumber.Substring(PhoneNumber.Length - 10, 10))
                End If
            End If
        End While
        InsertReader.Close()
        o1.Close()
        o1.Dispose()

        newXlApp.DisplayAlerts = False
        Dim CommonNoQuery As String
        Dim TotalNumbers As Integer = 0
        Dim queryString_b As String = "Select DISTINCT(bParty) from bParty order by bParty"
        Dim bPartyOthersCommand As SqlCommand
        Dim bPartyOthersReader As SqlDataReader
        Dim queryString_a As String = Nothing
        Dim queryString_Full_b As String = Nothing
        Dim oreader_a As OleDb.OleDbDataReader
        Dim oreader_b As OleDb.OleDbDataReader
        Dim oreader_full_b As OleDb.OleDbDataReader
        Dim checkedNumber As Integer = 0
        Dim totalfields As Integer = 0
        Dim IndexOfRow As Integer = 2
        Dim PreNumber As Long = 0
        Dim length As Integer = 0
        Dim IsSetFieldsName As Boolean = False
        Call ConnectionOpen()
        CommonNoQuery = "Select DISTINCT(bParty) from bParty order by bParty"
        bPartyOthersCommand = New SqlCommand(CommonNoQuery, TargetConnection)
        bPartyOthersReader = bPartyOthersCommand.ExecuteReader()
        While bPartyOthersReader.Read()
            TotalNumbers += 1
            lbTotalNos.Text = TotalNumbers
            lbTotalNos.Refresh()
        End While

        bPartyOthersReader.Close()

        bPartyOthersCommand = New SqlCommand(CommonNoQuery, TargetConnection)
        bPartyOthersReader = bPartyOthersCommand.ExecuteReader()
        Dim IsbpartyAdded As Boolean = False
        Dim NextNumber As Long = Nothing
        Dim NextNumbertxt As String = Nothing
        Dim isNumberFound As Integer = False
        Dim a As String = Nothing
        Dim b As String = Nothing
        Dim grpTime As String = Nothing
        Dim grpDate As String = Nothing
        Dim Duration As String = Nothing
        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        GroupCon = New OleDb.OleDbConnection(ConnectionString)
        GroupCon.Open()
        While bPartyOthersReader.Read()
            If IsDBNull(bPartyOthersReader(0)) = False Then
                isNumberFound = False
                'For k As Integer = 0 To NumberOfFiles - 1
                '    OnlyFileNames(k) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(k))
                'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
                'GroupCon = New OleDb.OleDbConnection(ConnectionString)
                'GroupCon.Open()
                'If NumberOfFiles = 1 Then
                queryString_a = "Select * from [" & SheetName & "] where a like '%" & bPartyOthersReader(0) & "'"
                'Else
                '    queryString_a = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "] where a like '%" & bPartyOthersReader(0) & "'"
                'End If
                ' queryString_a = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "] where a like '%" & bPartyOthersReader(0) & "'"
                ocmd_a = New OleDb.OleDbCommand(queryString_a, GroupCon)
                oreader_a = ocmd_a.ExecuteReader()
                ' isNumberFound = False
                While oreader_a.Read
                    If IsDBNull(oreader_a(0)) = False Then
                        isNumberFound = True
                        If IsDBNull(oreader_a("a")) = False Then
                            a = oreader_a("a")
                        Else
                            a = "NULL"
                        End If
                        If IsDBNull(oreader_a("b")) = False Then
                            b = oreader_a("b")
                        Else
                            b = "NULL"
                        End If
                        If IsDBNull(oreader_a("Time")) = False Then
                            grpTime = oreader_a("Time")
                        Else
                            grpTime = "NULL"
                        End If
                        If IsDBNull(oreader_a("Date")) = False Then
                            grpDate = oreader_a("Date")
                        Else
                            grpDate = "NULL"
                        End If
                        If IsDBNull(oreader_a("Duration")) = False Then
                            Duration = oreader_a("Duration")
                        Else
                            Duration = "NULL"
                        End If

                        If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then

                            For j As Integer = 1 To NumberOfFields

                                newXlWorkSheet.Cells(IndexOfRow, j) = oreader_a(j - 1)

                            Next
                            lbFound.Text = IndexOfRow - 1
                            IndexOfRow = IndexOfRow + 1
                        End If
                    End If
                End While
                oreader_a.Close()
                'oreader_b.Close()
                ocmd_a.Dispose()
                'ocmd_b.Dispose()
                'GroupCon.Close()
                'GroupCon.Dispose()
                '  Next
                If isNumberFound = True Then
                    'For m As Integer = 0 To NumberOfFiles - 1
                    '    OnlyFileNames(m) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(m))
                    'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
                    'GroupCon = New OleDb.OleDbConnection(ConnectionString)
                    'GroupCon.Open()
                    ' If NumberOfFiles = 1 Then
                    queryString_b = "Select * from [" & SheetName & "] Where b like '%" & bPartyOthersReader(0) & "'"
                    'Else
                    '    queryString_b = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "] Where b like '%" & bPartyOthersReader(0) & "'"
                    'End If
                    'queryString_b = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI from [" & SheetName1 & "] Where b like '%" & bPartyOthersReader(0) & "'"
                    ocmd_b = New OleDb.OleDbCommand(queryString_b, GroupCon)
                    oreader_b = ocmd_b.ExecuteReader()
                    While oreader_b.Read
                        If IsDBNull(oreader_b("a")) = False Then
                            a = oreader_b("a")
                        Else
                            a = "NULL"
                        End If
                        If IsDBNull(oreader_b("b")) = False Then
                            b = oreader_b("b")
                        Else
                            b = "NULL"
                        End If
                        If IsDBNull(oreader_b("Time")) = False Then
                            grpTime = oreader_b("Time")
                        Else
                            grpTime = "NULL"
                        End If
                        If IsDBNull(oreader_b("Date")) = False Then
                            grpDate = oreader_b("Date")
                        Else
                            grpDate = "NULL"
                        End If
                        If IsDBNull(oreader_b("Duration")) = False Then
                            Duration = oreader_b("Duration")
                        Else
                            Duration = "NULL"
                        End If

                        If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                            For j As Integer = 1 To NumberOfFields
                                newXlWorkSheet.Cells(IndexOfRow, j) = oreader_b(j - 1)
                            Next
                            lbFound.Text = IndexOfRow - 1
                            lbFound.Refresh()
                            IndexOfRow = IndexOfRow + 1
                        End If
                    End While
                    'oreader_a.Close()
                    oreader_b.Close()
                    'ocmd_a.Dispose()
                    ocmd_b.Dispose()
                    'GroupCon.Close()
                    'GroupCon.Dispose()
                    'Next
                End If
            End If
            checkedNumber += 1
            lbCheckedNos.Text = checkedNumber
            lbCheckedNos.Refresh()
        End While
        ' fileNameAndPath = fileNameAndPath.Insert(fileNameAndPath.IndexOf(".") - 1, " Groups")
        TargetConnection.Close()
        TargetConnection.Dispose()
        GroupCon.Close()
        GroupCon.Dispose()
        Try
            newXlWorkSheet.Columns.AutoFit()
        Catch ex As Exception

        End Try
        newXlApp.DisplayAlerts = False
        Try
            ' If NumberOfFiles = 1 Then
            Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(AllCommonFileNames(0))
            Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(0))
            Dim JustExtention As String = System.IO.Path.GetExtension(AllCommonFileNames(0))
            newXlWorkSheet.SaveAs(AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups"))

            'If cbFind_b_in_a.Checked = False Or isFromGroupPanel = True Then
            '    MsgBox("File has been created: " & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups"), MsgBoxStyle.Information)
            'End If
            'isFromGroupPanel = False
            'Else
            'newXlWorkSheet.SaveAs(SaveFilePath)
            'MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
            'End If

            'MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Please close the file" & SaveFilePath, MsgBoxStyle.Information)
        End Try


        newXlWorkbook.Close()
        newXlWorkbooks = Nothing
        newXlApp.Quit()
        newXlApp = Nothing
        If chkLimitedActivity.Checked = False Then
            If cbFind_b_in_a.Checked = False Or isFromGroupPanel = True Then
                MsgBox("File has been created: " & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups"), MsgBoxStyle.Information)
            End If
            isFromGroupPanel = False
            Exit Sub
        End If
        lbTotalNos.Text = "Initializing..."
        lbCheckedNos.Text = "Initializing..."
        lbFound.Text = "0"
        Call CreateNewExcelFile()
        Dim ColumnsHeaders As String = Nothing
        o.Open()
        ocmd1 = New OleDb.OleDbCommand(queryString_bOne, o)
        oreader_bone = ocmd1.ExecuteReader()
        NumberOfFields = oreader_bone.FieldCount
        ColNumber = 1
        For j As Integer = 1 To NumberOfFields
            newXlWorkSheet.Cells(1, j) = oreader_bone.GetName(j - 1).ToString
            If j > 1 Then
                ColumnsHeaders = ColumnsHeaders & ", [" & oreader_bone.GetName(j - 1).ToString & "] nvarchar(255)"
            ElseIf j = 1 Then
                ColumnsHeaders = "[" & oreader_bone.GetName(j - 1).ToString & "] nvarchar(255)"
            End If
            If newXlWorkSheet.Cells(1, j).value = "a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "CNIC_a" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "CNIC_b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "Time" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "hh:mm:ss AM/PM"
            ElseIf newXlWorkSheet.Cells(1, j).value = "Date" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "dd-mm-yy"
            ElseIf newXlWorkSheet.Cells(1, j).value = "b" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            ElseIf newXlWorkSheet.Cells(1, j).value = "IMEI" Then
                ColName = Chr(65 + ColNumber - 1)
                ColRange = ColName & ":" & ColName
                newXlWorkSheet.Range(ColRange).NumberFormat = "0"
            End If


            ColNumber = ColNumber + 1
        Next


        Call CreateTblLimimtedActivity("tblLimitedActivity", ColumnsHeaders)
        ColumnsHeaders = Nothing
        TotalNumbers = 0
        While oreader_bone.Read
            ColumnsHeaders = Nothing
            For j As Integer = 1 To NumberOfFields
                If j > 1 Then
                    If IsDBNull(oreader_bone(j - 1)) = False Then
                        ColumnsHeaders = ColumnsHeaders & ", '" & oreader_bone(j - 1).ToString & "'"
                    Else
                        ColumnsHeaders = ColumnsHeaders & ", 'NA'"
                    End If
                ElseIf j = 1 Then
                    If IsDBNull(oreader_bone(j - 1)) = False Then
                        ColumnsHeaders = "'" & oreader_bone(j - 1).ToString & "'"
                    Else
                        ColumnsHeaders = "'NA'"
                    End If
                End If
            Next
            Call InsertAnalyzedToSQL("tblLimitedActivity", ColumnsHeaders)
            TotalNumbers += 1
            lbTotalNos.Text = TotalNumbers
            lbTotalNos.Refresh()
        End While

        oreader_bone.Close()
        'ocmd1.Dispose()
        o.Close()
        'o.Dispose()
        queryString_bOne = "Select a,b from [" & SheetName & "]"

        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups") & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        o = New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        ocmd1 = New OleDb.OleDbCommand(queryString_bOne, o)
        oreader_bone = ocmd1.ExecuteReader()
        NumberOfFields = oreader_bone.FieldCount
        TotalNumbers = 0
        While oreader_bone.Read
            If IsDBNull(oreader_bone(0)) = False And IsDBNull(oreader_bone(1)) = False Then
                Call RmvGrpFromAnalyzed(oreader_bone(0).ToString, oreader_bone(1).ToString, "tblLimitedActivity")
            ElseIf IsDBNull(oreader_bone(0)) = False And IsDBNull(oreader_bone(1)) = True Then
                Call RmvGrpFromAnalyzed(oreader_bone(0).ToString, "NA", "tblLimitedActivity")
            ElseIf IsDBNull(oreader_bone(0)) = True And IsDBNull(oreader_bone(1)) = False Then
                Call RmvGrpFromAnalyzed("NA", oreader_bone(1).ToString, "tblLimitedActivity")
            End If
            TotalNumbers += 1
            lbCheckedNos.Text = TotalNumbers
            lbCheckedNos.Refresh()
        End While
        oreader_bone.Close()
        ocmd1.Dispose()
        o.Close()
        o.Dispose()
        TotalNumbers = 0
        Call ConnectionOpen()
        QueryString = "Select * from tblLimitedActivity"
        bPartyOthersCommand = New SqlCommand(QueryString, TargetConnection)
        bPartyOthersReader = bPartyOthersCommand.ExecuteReader()
        NumberOfFields = bPartyOthersReader.FieldCount
        IndexOfRow = 2
        While bPartyOthersReader.Read
            For j As Integer = 1 To NumberOfFields
                If bPartyOthersReader(j - 1) <> "NA" Then
                    newXlWorkSheet.Cells(IndexOfRow, j) = bPartyOthersReader(j - 1)
                Else
                    newXlWorkSheet.Cells(IndexOfRow, j) = ""
                End If
            Next
            TotalNumbers += 1
            lbFound.Text = TotalNumbers
            lbFound.Refresh()
            IndexOfRow = IndexOfRow + 1
        End While
        Try
            ' If NumberOfFiles = 1 Then
            newXlWorkSheet.Columns.AutoFit()
            newXlApp.DisplayAlerts = False
            Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(AllCommonFileNames(0))
            Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(0))
            Dim JustExtention As String = System.IO.Path.GetExtension(AllCommonFileNames(0))
            newXlWorkSheet.SaveAs(AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " LimitedActivity"))

            'If cbFind_b_in_a.Checked = False Or isFromGroupPanel = True Then
            '    MsgBox("File has been created: " & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups"), MsgBoxStyle.Information)
            'End If
            'isFromGroupPanel = False
            'Else
            'newXlWorkSheet.SaveAs(SaveFilePath)
            'MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
            'End If

            'MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Please close the file" & SaveFilePath, MsgBoxStyle.Information)
        End Try
        newXlWorkbook.Close()
        newXlWorkbooks = Nothing
        newXlApp.Quit()
        newXlApp = Nothing
        If cbFind_b_in_a.Checked = False Or isFromGroupPanel = True Then
            MsgBox("File has been created: " & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " Groups") & vbCrLf & "File has been created: " & AllCommonFileNames(0).Insert(AllCommonFileNames(0).IndexOf(".") - 1, " LimitedActivity"), MsgBoxStyle.Information)
        End If
        isFromGroupPanel = False

    End Sub
    Dim QueryString_b As String
    Sub RmvGrpFromAnalyzed(ByVal aParty As String, ByVal bParty As String, ByVal tblName As String)
        Call ConnectionOpen()
        'If aParty And bParty Then
        QueryString = "Delete From " & tblName & " Where " & "a=" & "'" & aParty & "'" & " AND " & "b=" & "'" & bParty & "'"
        'Else
        '    QueryString = "Delete From " & tblName & " Where " & "a=" & "'" & aParty & "'" & " AND " & "b=" & "''"
        'QueryString = "Delete From " & tblName & " Where " & "a LIKE " & "'" & aParty.Substring(aParty.Length - 11, 11) & " AND " & "b LIKE " & "'" & bParty.Substring(bParty.Length - 11, 11)
        '    QueryString_b = "Delete From " & tblName & " Where " & "b LIKE " & "'%" & aParty.Substring(aParty.Length - 10, 10) & "'"
        'ElseIf aParty.Length >= 10 And bParty.Length < 10 Then
        '    QueryString = "Delete From " & tblName & " Where " & "a=" & aParty & " AND " & "b=" & bParty
        '    QueryString_b = "Delete From " & tblName & " Where " & "b LIKE " & "'%" & aParty.Substring(aParty.Length - 10, 10) & "'"
        'ElseIf aParty.Length < 10 And bParty.Length < 10 Then
        '    QueryString = "Delete From " & tblName & " Where " & "a=" & aParty & " AND " & "b= " & bParty
        '    QueryString_b = "Delete From " & tblName & " Where " & "b=" & aParty
        'ElseIf aParty.Length < 10 And bParty.Length >= 10 Then
        '    QueryString = "Delete From " & tblName & " Where " & "a=" & aParty & " AND " & "b=" & bParty
        '    QueryString_b = "Delete From " & tblName & " Where " & "b=" & aParty
        ' End If
        CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
        Try
            CreateTbCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(QueryString, MsgBoxStyle.Information)
        End Try
        'CreateTbCommand = New SqlCommand(QueryString_b, TargetConnection)
        'Try
        '    CreateTbCommand.ExecuteNonQuery()
        'Catch ex As Exception
        '    MsgBox(QueryString, MsgBoxStyle.Information)
        'End Try
        TargetConnection.Close()
        TargetConnection.Dispose()
        CreateTbCommand.Dispose()
    End Sub

    Function IsDuplicate(ByVal tblName As String, ByVal a As String, ByVal b As String, ByVal grpTime As String, ByVal grpDate As String, ByVal Duration As String) As Boolean
        Dim IsDuplication As Boolean = False
        Call ConnectionOpenDelDupli()
        Dim InsertIMEICommand As SqlCommand
        Dim CheckReader As SqlDataReader
        Dim CheckQuery As String = "Select * From [" & tblName & "] Where a = '" & a & "' AND b = '" & b & "' AND grpTime = '" & grpTime & "' AND grpDate = '" & grpDate & "' AND Duration = '" & Duration & "'"
        InsertIMEICommand = New SqlCommand(CheckQuery, TargetConnectionDelDupli)
        CheckReader = InsertIMEICommand.ExecuteReader
        While CheckReader.Read
            IsDuplication = True
            Exit While
        End While
        CheckReader.Close()
        InsertIMEICommand.Dispose()
        If IsDuplication = False Then
            InsertQuery = "INSERT INTO [" & tblName & "] VALUES ('" & a & "','" & b & "','" & grpTime & "','" & grpDate & "','" & Duration & "')"
            'InsertQuery = "INSERT INTO [" & tblName & "] VALUES (" & ColumnsValues & ")"
            InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnectionDelDupli)
            Try
                InsertIMEICommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error in insert", MsgBoxStyle.Information)
            End Try
        End If
        TargetConnectionDelDupli.Close()
        TargetConnectionDelDupli.Dispose()
        InsertIMEICommand.Dispose()
        Return IsDuplication
    End Function
    Dim TblCheckDuplication As String = "TblCheckDuplication"
    Sub CreateTblDuplicateCheck(ByVal tblName As String)
        Try
            Call ConnectionOpenDelDupli()
            QueryString = "IF OBJECT_ID('dbo." & tblName & "') IS NOT NULL DROP TABLE " & tblName & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnectionDelDupli)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnectionDelDupli.Close()
        Catch ex As Exception
            If TargetConnectionDelDupli.State <> ConnectionState.Closed Then
                TargetConnectionDelDupli.Close()
            End If
        End Try
        Call ConnectionOpenDelDupli()
        Try
            QueryString = "CREATE TABLE " & tblName & "(a varchar(25), b varchar(25), grpTime varchar(50), grpDate varchar(50), Duration varchar(16))"
            'QueryString = "CREATE TABLE " & tblName & "(" & ColumnsHeaders & ")"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnectionDelDupli)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnectionDelDupli.Close()
            TargetConnectionDelDupli.Dispose()
            CreateTbCommand.Dispose()
        Catch ex As Exception

        End Try
    End Sub
    Sub DropTbl(ByVal tblName As String)
        Try
            Call ConnectionOpenDelDupli()
            QueryString = "IF OBJECT_ID('dbo." & tblName & "') IS NOT NULL DROP TABLE " & tblName & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnectionDelDupli)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnectionDelDupli.Close()
        Catch ex As Exception
            If TargetConnectionDelDupli.State <> ConnectionState.Closed Then
                TargetConnectionDelDupli.Close()
            End If
        End Try
    End Sub
    Sub CreateTblLimimtedActivity(ByVal tblName As String, ByVal ColumnsHeaders As String)
        Try
            Call ConnectionOpen()
            QueryString = "IF OBJECT_ID('dbo." & tblName & "') IS NOT NULL DROP TABLE " & tblName & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
        Catch ex As Exception
            If TargetConnection.State <> ConnectionState.Closed Then
                TargetConnection.Close()
            End If
        End Try
        Call ConnectionOpen()
        Try
            'QueryString = "CREATE TABLE " & tblName & "(a varchar(16),CNIC_a text, b varchar(16), CNIC_b text, Time datetime, Date datetime, [Call Type] varchar(15), Duration varchar(6),[Cell ID] varchar(25), IMEI varchar(25), IMSI varchar(25), Site text)"
            QueryString = "CREATE TABLE " & tblName & "(" & ColumnsHeaders & ")"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
            TargetConnection.Dispose()
            CreateTbCommand.Dispose()
        Catch ex As Exception

        End Try
    End Sub
    Sub InsertAnalyzedToSQL(ByVal tblName As String, ByVal ColumnsValues As String)
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        'InsertQuery = "INSERT INTO [" & tblName & "] VALUES ('" & a & "','" & CNIC_a & "','" & b & "','" & CNIC_b & "','" & Time & "','" & GrpDate & "','" & CallType & "','" & Duration & "','" & CellID & "','" & IMEI & "','" & IMSI & "','" & Site & "')"
        InsertQuery = "INSERT INTO [" & tblName & "] VALUES (" & ColumnsValues & ")"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        Try
            InsertIMEICommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error in insert" & vbCrLf & Err.Description, MsgBoxStyle.Information)
        End Try

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub Find_b_in_a(ByVal fileNameAndPath As String)
        'Dim xlApp As Excel.Application = Nothing
        'Dim xlWorkbooks As Excel.Workbooks = Nothing
        'Dim xlWorkbook As Excel.Workbook = Nothing
        'Dim xlSheet As Excel.Worksheet = Nothing
        'Dim xlSheet2 As Excel.Worksheet = Nothing
        'xlApp = New Excel.Application
        Dim SheetName1 As String = "Sheet1$"
        Dim SheetName2 As String = "Sheet1$"
        'xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        'xlApp.DisplayAlerts = False
        Dim NumberOfFields As Integer = 0
        Dim filePath As String = fileNameAndPath
        fileNameAndPath = fileNameAndPath.Insert(fileNameAndPath.IndexOf(".") - 1, " Groups")
        ' xlWorkbook = xlApp.Workbooks.Open(lbfilepath.Text, , , , , , , , , True, False)
        Call CreateNewExcelFile()
        'xlSheet = xlWorkbook.Worksheets("Sheet1")
        'xlSheet2 = xlWorkbook.Worksheets("Sheet2")
        newXlApp.DisplayAlerts = False
        Dim queryString_b As String = "Select * from [" & SheetName1 & "] order by b"
        Dim queryString_a As String = Nothing
        Dim queryString_Full_b As String = Nothing
        Dim oreader_a As OleDb.OleDbDataReader
        Dim oreader_full_b As OleDb.OleDbDataReader
        Dim ocmd_a As OleDbCommand
        Dim ocmd_Full_b As OleDbCommand
        Dim checkedNumber As Integer = 0
        Dim totalfields As Integer = 0
        Dim IndexOfRow As Integer = 2
        Dim PreNumber As Long = 0
        Dim length As Integer = 0
        Dim IsSetFieldsName As Boolean = False
        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        Dim IsbpartyAdded As Boolean = False
        Dim NextNumber As Long = Nothing
        Dim NextNumbertxt As String = Nothing
        Dim ocmd1 As New OleDb.OleDbCommand(queryString_b, o)
        'ocmd1.ExecuteNonQuery()
        ' Dim oreader As OleDb.OleDbDataReader
        Dim oreader_b As OleDb.OleDbDataReader
        oreader_b = ocmd1.ExecuteReader()
        totalfields = oreader_b.FieldCount
        For j As Integer = 1 To totalfields
            newXlWorkSheet.Cells(1, j) = oreader_b.GetName(j - 1).ToString
        Next
        Dim isNumberFound As Integer = False
        While oreader_b.Read()
            If IsDBNull(oreader_b("b")) = False Then
                NextNumber = oreader_b("b")
                NextNumbertxt = NextNumber
                length = NextNumbertxt.Length
                If length >= 10 Then
                    If PreNumber <> NextNumber Then
                        isNumberFound = False
                        queryString_a = "Select * from [" & SheetName1 & "] where a =" & oreader_b("b") & ""
                        ocmd_a = New OleDb.OleDbCommand(queryString_a, o)
                        oreader_a = ocmd_a.ExecuteReader()
                        IsbpartyAdded = False
                        While oreader_a.Read()
                            isNumberFound = True
                            For j As Integer = 1 To totalfields
                                newXlWorkSheet.Cells(IndexOfRow, j) = oreader_a(j - 1)
                            Next
                            '    IsSetFieldsName = True
                            lbFound.Text = IndexOfRow - 1
                            IndexOfRow = IndexOfRow + 1
                            If IsbpartyAdded = False Then
                                For i As Integer = 1 To totalfields
                                    newXlWorkSheet.Cells(IndexOfRow, i) = oreader_b(i - 1)
                                    IsbpartyAdded = True
                                Next
                                lbFound.Text = IndexOfRow - 1
                                IndexOfRow = IndexOfRow + 1
                            End If
                        End While
                        oreader_a.Close()
                    ElseIf isNumberFound = True And PreNumber = oreader_b("b") Then
                        For k As Integer = 1 To totalfields
                            newXlWorkSheet.Cells(IndexOfRow, k) = oreader_b(k - 1)
                        Next
                        lbFound.Text = IndexOfRow - 1
                        IndexOfRow = IndexOfRow + 1
                    End If
                    PreNumber = oreader_b("b")
                End If
            End If
            checkedNumber += 1
            lbCheckedNos.Text = checkedNumber
        End While
        oreader_b.Close()
        o.Close()
        o.Dispose()
        Dim DirectoryPath As String = System.IO.Path.GetDirectoryName(lbfilepath.Text)
        Dim JustFileName As String = System.IO.Path.GetFileNameWithoutExtension(lbfilepath.Text)
        Dim JustExtention As String = System.IO.Path.GetExtension(lbfilepath.Text)
        newXlApp.DisplayAlerts = False
        Try
            newXlWorkSheet.SaveAs(fileNameAndPath)
        Catch ex As Exception
            MsgBox("Please close the file" & JustFileName & "(AnalyzedGroups)" & JustExtention, MsgBoxStyle.Information)
        End Try
        newXlWorkbook.Close()
        newXlWorkbooks = Nothing
        newXlApp.Quit()
        newXlApp = Nothing
    End Sub
    Dim PathAndFileName As String = Nothing
    Dim OnlyFileName As String = Nothing
    Public DocTitle As String
    Public PathForDoc As String

    Private Sub btn_BrowsBTS_Click(sender As System.Object, e As System.EventArgs) Handles btn_BrowsBTS.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        Dim FileExtention As String = Nothing
        OpenFileDialog1.Filter = "Excel Files |*.xlsx;*.xls"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            PathAndFileName = OpenFileDialog1.FileName
            PathForDoc = OpenFileDialog1.FileName
            OnlyFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            DocTitle = OnlyFileName
            AllCommonFileNames = OpenFileDialog1.FileNames()
            NumberOfFiles = AllCommonFileNames.Length()
            If FileExtention = ".xlsx" Then
                PathForDoc = PathForDoc.Substring(0, PathForDoc.Length - 5)
            Else
                PathForDoc = PathForDoc.Substring(0, PathForDoc.Length - 4)
            End If
            txt_BTS_Path.Text = PathAndFileName
            btn_CropBTS.Enabled = True
        Else
            Exit Sub
        End If
        '        Call ChangeTimeFormat("@")
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'Call createTempTable()
        For i = 1 To 10
            ' Call InsertIMEI("1234567589" & i)
        Next
        If Button1.Text = "Stop" Then
            Button1.Text = "Close"
            StopAnalyze = True
        Else
            frm_Spy_Tech.Enabled = True
            Me.Close()
        End If

    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Top = frm_Spy_Tech.Panel1.Height + 30
        frm_Spy_Tech.Enabled = False
        cbFind_b_in_a.Checked = False
        'Call ConditionalDisplayBTS()
    End Sub
    Dim FolderPath As String = Nothing
    Dim SaveGroupExtention As String = Nothing
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'Dim OpenFileDialog1 As New OpenFileDialog
        'Dim OnlyFileName As String = Nothing
        'Dim FileExtention As String = Nothing
        'Dim PathForDoc As String = Nothing
        'Dim DocTitle As String = Nothing
        'OpenFileDialog1.Filter = "Excel Files |*.xlsx;*.xls"
        'If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
        '    PathAndFileName = OpenFileDialog1.FileName
        '    PathForDoc = OpenFileDialog1.FileName
        '    OnlyFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
        '    FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
        '    DocTitle = OnlyFileName
        '    If FileExtention = ".xlsx" Then
        '        PathForDoc = PathForDoc.Substring(0, PathForDoc.Length - 5)
        '    Else
        '        PathForDoc = PathForDoc.Substring(0, PathForDoc.Length - 4)
        '    End If
        '    lbfilepath.Text = PathAndFileName

        'Else
        '    Exit Sub
        'End If
        ' Call Find_b_in_a()
        Dim OpenFileDialog1 As New OpenFileDialog

        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Excel Files|*.xlsx;*.xls"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            NumberOfFiles = AllFileNames.Length()
            lbfilepath.Text = NumberOfFiles & " files have been selected"
            'If NumberOfFiles < 2 Then
            '    MsgBox("Please select more than one files to compare", vbOKOnly)
            '    btnIMEIsComparison.Enabled = False
            '    Exit Sub
            'Else
            '    btnIMEIsComparison.Enabled = True
            'End If
            AllCommonFileNames = OpenFileDialog1.FileNames()
            FolderPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            SaveGroupExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            OnlyFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            If FileExtention = ".xlsx" Then
                OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            Else
                OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            End If
            'lbfilepath.Text = PathAndFileName
            'lbIMEIsNames.Text = PathAndFileName
        Else
            Exit Sub
        End If
    End Sub
    Dim NumberOfFiles As Integer
    Dim isFromGroupPanel As Boolean = False
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' Call FindGroups()
        If lbfilepath.Text <> "" Then
            ' Try
            If NumberOfFiles = 1 Then
                ' Call Find_b_in_a(AllCommonFileNames(0))
                isFromGroupPanel = True
                Call FindGroupsOfOne()

                ' MsgBox(lbfilepath.Text + "has been created", MsgBoxStyle.Information)
            ElseIf NumberOfFiles > 1 Then
                Call FindGroups()
            End If
            'Catch ex As Exception
            '    MsgBox("Error", MsgBoxStyle.Information)
            'End Try
            Call DropTbl(TblCheckDuplication)
        End If
    End Sub

    Public AllCommonFileNames() As String = Nothing
    Public OnlyFilesPath As String = Nothing


    Private Sub btnBrowseBTSs_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowseBTSs.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Excel Files|*.xlsx;*.xls"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
            lbIMEIsNames.Text = NumberOfFiles & " files have been selected"
            If NumberOfFiles < 2 Then
                MsgBox("Please select more than one files to compare", vbOKOnly)
                btnIMEIsComparison.Enabled = False
                Exit Sub
            Else
                btnIMEIsComparison.Enabled = True
            End If
            AllCommonFileNames = OpenFileDialog1.FileNames()

            OnlyFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            OnlyFilesPath = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
            If FileExtention = ".xlsx" Then
                OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 5)
            Else
                OnlyFileName = OnlyFileName.Substring(0, OnlyFileName.Length - 4)
            End If
            'lbIMEIsNames.Text = PathAndFileName
        Else
            Exit Sub
        End If
    End Sub
    Dim IMEIOthersCommand As SqlCommand
    Dim IMEIOthersReader As SqlDataReader
    Private Sub btnIMEIsComparison_Click(sender As System.Object, e As System.EventArgs) Handles btnIMEIsComparison.Click
        'Creating word Document
        Me.Cursor = Cursors.WaitCursor
        DT_PhoneNumber.Clear()
        lbTotalIMEIs.Text = "Initializing..."
        lbCheckedIMEI.Text = "Initializing..."
        lbFoundIMEIs.Text = "Wait..."
        ' DVG_SpyTech.Rows.Clear()
        'Lstbx_CNIC_PhoneNumbers.Items.Clear()
        'prgbar_Common_Links.Minimum = 0
        'prgbar_Common_Links.Maximum = AllCommonFileNames.Length * 3 + 2
        'prgbar_Common_Links.Value = 0
        'prgbar_Common_Links.Visible = True
        Dim WordApp As New Word.Application()
        Dim doc As New Word.Document()
        doc = WordApp.Documents.Add()
        Dim CommonNoTable As Word.Table
        With doc.Range
            .InsertAfter("BTS IMEI Comparison Report")
            .InsertParagraphAfter()
            .InsertAfter("Date:  " & Date.Today & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Time:  " & Date.Now.ToLongTimeString)
            .InsertParagraphAfter()
            .InsertParagraphAfter()
        End With
        With doc.PageSetup
            .PaperSize = Word.WdPaperSize.wdPaperA4
            .LeftMargin = 20
            .RightMargin = 20
            .TopMargin = 30
            .BottomMargin = 30
        End With


        Dim SelRange As Word.Range
        SelRange = doc.Paragraphs.Item(1).Range
        SelRange.Font.Size = 14
        SelRange.Font.Bold = True
        SelRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        SelRange = doc.Paragraphs.Item(2).Range
        SelRange.Font.Size = 12
        SelRange.Font.Bold = True
        SelRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter


        Dim TotalNumberOfCDRs As Integer = AllCommonFileNames.Length()
        Dim OnlyFileNames(TotalNumberOfCDRs) As String
        If TotalNumberOfCDRs > 4 Then
            doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
        End If
        'prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
        Dim NumberOfCommons(TotalNumberOfCDRs) As String
        'Creating Table in Word Docurment
        CommonNoTable = doc.Range.Tables.Add(doc.Bookmarks.Item("\endofdoc").Range, 1, TotalNumberOfCDRs + 2)
        CommonNoTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
        CommonNoTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        CommonNoTable.Borders.InsideColor = Word.WdColor.wdColorBlack
        CommonNoTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        CommonNoTable.Range.Font.Size = 9
        CommonNoTable.Columns.AutoFit()


        CommonNoTable.Range.ParagraphFormat.LineSpacing = 1
        CommonNoTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast
        CommonNoTable.Range.ParagraphFormat.SpaceBefore = 0
        CommonNoTable.Range.ParagraphFormat.SpaceAfter = 0
        CommonNoTable.Style = "Light Grid"
        CommonNoTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
        CommonNoTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        CommonNoTable.Rows(1).Cells(1).Range.Text = "Sr.NO"

        CommonNoTable.Rows(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
        CommonNoTable.Rows(1).Cells(2).Range.Text = "IMEI" '& vbCrLf & "Numbers"
        CommonNoTable.Rows(1).Range.Font.Bold = True
        'CommonNoTable.Cell(1, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20
        'CommonNoTable.Cell(1, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20
        'CommonNoTable.Rows.Item(1).Range.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        'CommonNoTable.Rows.Item(1).Range.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic
        'CommonNoTable.Rows.Item(1).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray30
        'CommonNoTable.Columns(2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10
        'CommonNoTable.Columns.AutoFit()
        'CommonNoTable.Range.ParagraphFormat.LeftIndent=
        ' prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
        Dim ColumnName1 As String = "IMEI"
        Dim ColumnName2 As String = "b"
        Dim SheetName As String = "bts$"
        Dim ConnectionString As String
        Dim Commandbts As SqlCommand
        Dim Readerbts As SqlDataReader
        Try
            Dim IMEI_Number As String

            'Create DataTable for results
            'OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            'OthersConnection.Open()
            'Dim TargetConnection As SqlConnection
            'TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            'TargetConnection.Open()
            'Dim Delquery As String = "DELETE FROM CommonNumbers"
            'OthersCommand = New SqlCommand(Delquery, OthersConnection)
            'OthersCommand.ExecuteNonQuery()
            'Dim CreateTablesQuery As String
            Dim TempTableName As String
            'Dim QueryInsert As String
            Call DropTempTable("IMEITable")

            Call createTempTable("IMEITable", True)
            For i As Integer = 0 To TotalNumberOfCDRs - 1
                Try
                    OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                    TempTableName = "_" & OnlyFileNames(i)
                    'CommonNoTable.Rows(1).Cells(i + 3).Range.Text = OnlyFileNames(i)
                    Call DropTempTable("[" & TempTableName & "]")
                    Call createTempTable(TempTableName, False)
                    'Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
                    'OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
                    'OthersCommand.ExecuteNonQuery()
                Catch ex1 As Exception

                End Try

            Next
            Dim queryString As String = Nothing
            For j As Integer = 0 To TotalNumberOfCDRs - 1
                ' prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
                OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
                TempTableName = "_" & OnlyFileNames(j)
                CommonNoTable.Rows(1).Cells(j + 3).Range.Text = OnlyFileNames(j)
                'CommonNoTable.Cell(1, j + 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20
                ' CreateTablesQuery = "CREATE TABLE [" & TempTableName & "] (PhoneNumber varchar(50) null)"
                'OthersCommand = New SqlCommand(CreateTablesQuery, OthersConnection)
                'OthersCommand.ExecuteNonQuery()
                ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
                Dim o As New OleDb.OleDbConnection(ConnectionString)
                o.Open()
                'populate table with data with 92
                queryString = "Select a, IMEI from [" & SheetName & "] GROUP BY a,IMEI"
                Dim InsertCommand As New OleDb.OleDbCommand(queryString, o)
                Dim InsertReader As OleDb.OleDbDataReader
                Try
                    InsertReader = InsertCommand.ExecuteReader()
                Catch ex As Exception
                    Call DropTempTable("IMEITable")
                    For i As Integer = 0 To TotalNumberOfCDRs - 1
                        Try
                            OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                            TempTableName = "_" & OnlyFileNames(i)
                            Call DropTempTable("[" & TempTableName & "]")
                        Catch ex1 As Exception

                        End Try

                    Next
                    Me.Cursor = Cursors.Default
                    MsgBox("The result has not been produced" & vbCrLf & "Please make sure sheet is 'bts' with columns name 'a' , 'b' and 'IMEI' of the file " & TempTableName.Substring(1), MsgBoxStyle.Information)
                    Exit Sub
                End Try


                While InsertReader.Read
                    If IsDBNull(InsertReader(1)) = False Then

                        If IMEI_Number <> InsertReader(1).ToString Then
                            IMEI_Number = InsertReader(1)
                            If IMEI_Number.Length > 14 Then
                                IMEI_Number = IMEI_Number.Substring(0, 14)
                            End If
                            Call InsertIMEI(IMEI_Number)
                            Call Insert_a_IMEI(InsertReader(0).ToString, IMEI_Number, TempTableName)
                        End If
                    End If
                End While
                InsertReader.Close()
                o.Close()
                o.Dispose()
            Next
            Dim TotalIMEIs As Long = 0
            Call ConnectionOpen()
            Dim CommonNoQuery As String = "SELECT DISTINCT(IMEI) FROM IMEITable" 'GROUP BY IMEI"

            IMEIOthersCommand = New SqlCommand(CommonNoQuery, TargetConnection)
            IMEIOthersReader = IMEIOthersCommand.ExecuteReader()
            While IMEIOthersReader.Read()
                TotalIMEIs += 1
            End While
            lbTotalIMEIs.Text = TotalIMEIs
            IMEIOthersReader.Close()

            IMEIOthersCommand = New SqlCommand(CommonNoQuery, TargetConnection)
            IMEIOthersReader = IMEIOthersCommand.ExecuteReader()
            Dim CheckedIMEIs As Integer = 0
            Dim IMEINumber As String = Nothing
            Dim NumberofRows As String = Nothing
            Dim IsFound As Boolean = False
            Dim RecCounter As Integer
            Dim RowIndex As Integer = 2
            Dim TargetQueryString As String
            Dim IsMoreThanOne As Boolean = False
            Dim PathAndFileName As String
            ' prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
            While IMEIOthersReader.Read()
                ' NumberOfCommons = Nothing
                RecCounter = 0
                IsMoreThanOne = False
                'TotalIMEIs = OthersReader(1)
                '   DT_CommonNumbers.Rows.Add()
                CommonNoTable.Rows.Add()
                'DT_CommonNumbers.Rows(RowIndex).Item(0) = OthersReader(0).ToString
                If IsDBNull(IMEIOthersReader(0)) = False Then

                    IMEINumber = IMEIOthersReader(0).ToString

                    For i As Integer = 0 To TotalNumberOfCDRs - 1

                        PathAndFileName = AllCommonFileNames(i)
                        OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                        ' Dim currentFileName = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                        TempTableName = "_" & OnlyFileNames(i)

                        TargetQueryString = "select Distinct(a)  from  [" & TempTableName & "]  WHERE IMEI = '" & IMEINumber & "'" 'GROUP BY IMEI"
                        'OthersConnection.Open()
                        'Dim TargetFileQueryString As String = "select PhoneNumber, count(*) as CountOf from  " & TempTableName & "  WHERE PhoneNumber LIKE '" & "%" & PhoneNumber & "%" & "'  GROUP BY PhoneNumber"
                        Call SearchIMEICon()
                        Commandbts = New SqlCommand(TargetQueryString, SearchIMEIConnection)

                        Readerbts = Commandbts.ExecuteReader
                        'TargetReader = TargetFileCommandbts.ExecuteReader()

                        IsMoreThanOne = False
                        While Readerbts.Read()
                            ' NumberOfCommons(i) = Readerbts(1).ToString
                            '           DT_CommonNumbers.Rows(RowIndex).Item(OnlyFileNames(i)) = NumberOfCommons(i)
                            If IsMoreThanOne = False Then
                                With CommonNoTable
                                    .Rows(RowIndex).Cells(1).Range.Text = RowIndex - 1
                                    .Rows(RowIndex).Cells(2).Range.Text = IMEINumber
                                    .Rows(RowIndex).Cells(i + 3).Range.Text = Readerbts(0)
                                End With
                                RecCounter = RecCounter + 1
                                IsMoreThanOne = True
                            End If
                            '  MsgBox("Number has been found in " + TargetReader(0).ToString + " " + TargetReader(1).ToString, vbOKOnly)
                        End While
                        If RecCounter > 1 Then

                        End If

                        Readerbts.Close()
                        SearchIMEIConnection.Close()
                        SearchIMEIConnection.Dispose()
                        Commandbts.Dispose()

                    Next
                    If RecCounter <= 1 Then
                        '      DT_CommonNumbers.Rows.RemoveAt(RowIndex)
                        CommonNoTable.Rows(RowIndex).Delete()
                    Else

                        RowIndex = RowIndex + 1
                        lbFoundIMEIs.Text = RowIndex - 2
                    End If
                End If
                CheckedIMEIs += 1
                lbCheckedIMEI.Text = CheckedIMEIs
            End While
            'NumbersNotFound = NumbersNotFound & vbCrLf & vbCrLf
            IMEIOthersReader.Close()
            TargetConnection.Close()
            TargetConnection.Dispose()
            'For i As Integer = 0 To TotalNumberOfCDRs - 1
            '    '      prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
            '    TempTableName = "_" & OnlyFileNames(i)
            '    Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
            '    OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
            '    OthersCommand.ExecuteNonQuery()
            'Next
            CommonNoTable.Columns.AutoFit()
        Catch ex As Exception
            For i As Integer = 0 To TotalNumberOfCDRs - 1
                Try
                    OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                    Dim TempTableName As String = "_" & OnlyFileNames(i)
                    Call DropTempTable("[" & TempTableName & "]")
                    'Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
                    'OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
                    'OthersCommand.ExecuteNonQuery()
                Catch ex1 As Exception
                    GoTo nextItration
                End Try
nextItration:
            Next
            Call DropTempTable("IMEITable")

            MsgBox("The result has not been produced" & vbCrLf & "Please make sure sheet is 'bts' with columns name a and b", vbOKOnly)
            ' prgbar_Common_Links.Visible = False
            'prgbar_Common_Links.Value = 0
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try

        Try
            WordApp.Options.SavePropertiesPrompt = False
            doc.SaveAs(OnlyFilesPath & "\BTSimeiComparison.docx")
            doc.Close(Word.WdSaveOptions.wdSaveChanges)

            'prgbar_Common_Links.Value = AllCommonFileNames.Length * 3 + 2
            'prgbar_Common_Links.Visible = False
            'prgbar_Common_Links.Value = 0
            For i As Integer = 0 To TotalNumberOfCDRs - 1
                Try
                    OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                    Dim TempTableName As String = "_" & OnlyFileNames(i)
                    Call DropTempTable("[" & TempTableName & "]")
                    'Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
                    'OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
                    'OthersCommand.ExecuteNonQuery()
                Catch ex1 As Exception
                    GoTo nextItration1
                End Try
nextItration1:
            Next
            Call DropTempTable("IMEITable")
            MsgBox("The File has been created: " & vbCrLf & OnlyFilesPath & "\BTSimeiComparison", vbOKOnly)
        Catch ex As Exception
            'Dim MSWord As New Word.Application
            'Dim WordDoc As New Word.Document
            'prgbar_Common_Links.Visible = False
            'prgbar_Common_Links.Value = 0
            If File.Exists(OnlyFilesPath & "\BTSimeiComparison.docx") Then
                doc.Close(Word.WdSaveOptions.wdSaveChanges)

            End If
            For i As Integer = 0 To TotalNumberOfCDRs - 1
                Try
                    OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                    Dim TempTableName As String = "_" & OnlyFileNames(i)
                    Call DropTempTable("[" & TempTableName & "]")
                    'Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
                    'OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
                    'OthersCommand.ExecuteNonQuery()
                Catch ex1 As Exception
                    GoTo nextItration2
                End Try
nextItration2:
            Next
            Call DropTempTable("IMEITable")
            'doc.SaveAs(OnlyFilesPath & "\CommonNumberReport.docx")
            'doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            'MsgBox("The File has been created: " & vbCrLf & OnlyFilesPath & "\CommonNumbersReport", vbOKOnly)
        End Try

        Me.Cursor = Cursors.Default
        'btn_Save_in_MSWord.Enabled = False
        'OthersConnection.Close()
    End Sub
    Dim TargetConnection As SqlConnection
    Dim TargetConnectionDelDupli As SqlConnection
    Dim SearchIMEIConnection As SqlConnection
    Dim OutOfLimitConnection As SqlConnection
    Dim ConnectionWithinLimit As SqlConnection
    Dim ConnectionUpdatingLimit As SqlConnection
    Dim QueryString As String
    Dim CreateTbCommand As SqlCommand
    Sub SearchIMEICon()
        SearchIMEIConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If SearchIMEIConnection.State <> ConnectionState.Open Then
            SearchIMEIConnection.Open()
        End If
    End Sub
    Sub OutOfLimmitConn()
        OutOfLimitConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If OutOfLimitConnection.State <> ConnectionState.Open Then
            OutOfLimitConnection.Open()
        End If
    End Sub
    Sub ConWithinLimmit()
        ConnectionWithinLimit = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If ConnectionWithinLimit.State <> ConnectionState.Open Then
            ConnectionWithinLimit.Open()
        End If
    End Sub
    Sub ConUpdatingInLimit()
        ConnectionUpdatingLimit = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If ConnectionUpdatingLimit.State <> ConnectionState.Open Then
            ConnectionUpdatingLimit.Open()
        End If
    End Sub
    Sub ConnectionOpenDelDupli()
        TargetConnectionDelDupli = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If TargetConnectionDelDupli.State = ConnectionState.Closed Then
            TargetConnectionDelDupli.Open()
        End If
    End Sub
    Sub ConnectionOpen()
        TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If TargetConnection.State = ConnectionState.Closed Then
            TargetConnection.Open()
        End If
    End Sub
    Sub InlimitConnection()
        InLimitSqlConn = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If InLimitSqlConn.State = ConnectionState.Closed Then
            InLimitSqlConn.Open()
        End If
    End Sub
    Dim OutLimitSqlConn As SqlConnection
    Dim OutLimitSqlCommand As SqlCommand
    Dim OutLimitSqlReader As SqlDataReader
    Sub OutlimitConnection()
        OutLimitSqlConn = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If OutLimitSqlConn.State = ConnectionState.Closed Then
            OutLimitSqlConn.Open()
        End If
    End Sub
    Sub Create_bParyt_Table(ByVal TableName As String)
        Call ConnectionOpen()
        QueryString = "CREATE TABLE " & TableName & "(bParty varchar(16))"
        CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
        CreateTbCommand.ExecuteNonQuery()
        TargetConnection.Close()
        TargetConnection.Dispose()
        CreateTbCommand.Dispose()
    End Sub

    Sub createTempTable(ByVal TableName As String, ByVal isIMEITable As Boolean)
        Try
            'Call DropTempTable()


            'If TargetConnection.State <> ConnectionState.Open Then
            '    TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            '    TargetConnection.Open()
            'End If
            Call ConnectionOpen()
            'QueryString = "IF OBJECT_ID('tempdb.#IMEITable') IS NOT NULL DROP TABLE #IMEITable"
            'OthersCommand = New SqlCommand(QueryString, TargetConnection)
            'OthersCommand.ExecuteNonQuery()
            'QueryString = "CREATE TABLE #IMEITable(IMEI varchar(16))"
            If isIMEITable = True Then
                QueryString = "CREATE TABLE " & TableName & "(IMEI varchar(150))"
                CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
                CreateTbCommand.ExecuteNonQuery()
                TargetConnection.Close()
                TargetConnection.Dispose()
                CreateTbCommand.Dispose()
            Else
                QueryString = "CREATE TABLE [" & TableName & "] (a varchar(100),IMEI varchar(150))"
                CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
                CreateTbCommand.ExecuteNonQuery()
                TargetConnection.Close()
                TargetConnection.Dispose()
                CreateTbCommand.Dispose()
            End If
            'TargetConnection.Close()
            'TargetConnection.Close()
        Catch ex As Exception

            'If TargetConnection.State <> ConnectionState.Closed Then
            '    TargetConnection.Close()
            'End If
            'TargetConnection.Close()
        End Try



    End Sub
    Sub DropTempTable(ByVal TableName As String)
        Try

            'TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            'TargetConnection.Open()
            Call ConnectionOpen()
            ' QueryString = "DROP TABLE #IMEITable"
            ' QueryString = "IF OBJECT_ID('tempdb.#IMEITable') IS NOT NULL DROP TABLE #IMEITable"

            QueryString = "IF OBJECT_ID('dbo." & TableName & "') IS NOT NULL DROP TABLE " & TableName & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
        Catch ex As Exception
            If TargetConnection.State <> ConnectionState.Closed Then
                TargetConnection.Close()
            End If
        End Try
    End Sub
    Dim InsertQuery As String
    Dim InsertIMEICommand As SqlCommand
    Sub InsertPhoneNumber(ByVal PhoneNumber As String)
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO bParty VALUES ('" & PhoneNumber & "')"
        '" & CNIC & "'"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        InsertIMEICommand.ExecuteNonQuery()

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub InsertIMEI(ByVal IMEI As String)
        'TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=tempdb;Trusted_Connection=True;")
        'TargetConnection.Open()
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO IMEITable VALUES ('" & IMEI & "')"
        '" & CNIC & "'"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        InsertIMEICommand.ExecuteNonQuery()

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub Insert_a_IMEI(ByVal a As String, ByVal IMEI As String, ByVal TableName As String)
        'TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=tempdb;Trusted_Connection=True;")
        'TargetConnection.Open()
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO [" & TableName & "] VALUES ('" & a & "','" & IMEI & "')"
        '" & CNIC & "'"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        InsertIMEICommand.ExecuteNonQuery()
        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub


    Private Sub btnBrowseGroupSheet_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowseGroupSheet.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim OnlyFileName As String = Nothing
        Dim FileExtention As String = Nothing
        Dim PathForDoc As String = Nothing
        Dim DocTitle As String = Nothing
        OpenFileDialog1.Filter = "Excel Files |*.xlsx;*.xls"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            PathAndFileName = OpenFileDialog1.FileName
            PathForDoc = OpenFileDialog1.FileName
            OnlyFileName = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
            FileExtention = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
            DocTitle = OnlyFileName
            If FileExtention = ".xlsx" Then
                PathForDoc = PathForDoc.Substring(0, PathForDoc.Length - 5)
            Else
                PathForDoc = PathForDoc.Substring(0, PathForDoc.Length - 4)
            End If
            lbGroupSheetPath.Text = PathAndFileName

        Else
            Exit Sub
        End If
    End Sub
    Dim layerCounter As Integer = 1
    Dim Layer_Name As String = "Layer"
    Dim tracingTable As String = "tblAngramTrace"
    Dim GroupCounter As Integer = 0
    Sub CreateAnagramExcelFile()
        newXlApp = New Excel.Application
        newXlWorkbook = newXlApp.Workbooks.Add()

        ' xlSheets = newXlWorkbook.Sheets
        'newXlWorkSheet = "bts"
        ''newXlWorkSheet = newXlWorkbook.Sheets("Sheet1")
        ''newXlWorkSheet.Name = "bts"
    End Sub
    Sub LayeringAnagram()
        Dim PageCounter As Integer = 0
        Dim vDirectoryPath As String = System.IO.Path.GetDirectoryName(lbGroupSheetPath.Text)
        Dim vFilePath As String = Nothing
        Dim is_VFileSave As Boolean = False
        Dim vFileCounter As Integer = 1
        Dim xlFilePath As String = Nothing
        GroupCounter = 0
        lbTotalNoOfaParty.Text = "Initializing..."
        lbCheckedaPraty.Text = "Initializing..."
        lbNoOfAnagram.Text = "Initializing.."
        Try
            Call CreateAnagramTable("tblAnagram")
        Catch ex As Exception
            MsgBox("Error in Creating Anagram Table", MsgBoxStyle.Information)
            Exit Sub
        End Try
        Try
            Call BTS_to_SQL()
        Catch ex As Exception
            MsgBox(Err.Description & "Error in BTS to SQL", MsgBoxStyle.Information)
            Exit Sub
        End Try

        lbTotalNoOfaParty.Text = RecordsCount()
        ' Dim vDirectoryPath As String = System.IO.Path.GetDirectoryName(lbGroupSheetPath.Text)
        If RecordsCount() > 0 Then
            Try
                'Creating visio document
                Call CreateVisioDoc()
                'creating excel file
                Call CreateAnagramExcelFile()
            Catch ex As Exception
                MsgBox(Err.Description & "Error in Creating Visio File and Document", MsgBoxStyle.Information)
                Exit Sub
            End Try
        End If

        While RecordsCount() >= 1
            If PageCounter = 25 Then
                vFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (PageCounter * (vFileCounter - 1)) + 1 & " to " & GroupCounter & ".vsd"
                xlFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (PageCounter * (vFileCounter - 1)) + 1 & " to " & GroupCounter & ".xlsx"
                Try
                    ' Delete the previous version of the file.
                    Kill(vFilePath)
                    Kill(xlFilePath)
                    newXlWorkSheet.Columns.AutoFit()
                Catch
                End Try
                vFileCounter += 1
                vDoc.SaveAs(vFilePath)
                newXlWorkbook.SaveAs(xlFilePath)
                newXlWorkbook.Close()
                is_VFileSave = True
                PageCounter = 0
                vDoc.Close()
                vDoc = vApp.Documents.Add("")
                newXlWorkbook = newXlApp.Workbooks.Add()
                vDoc.PaperSize = Visio.VisPaperSizes.visPaperSizeLegal
                'vApp.Visible = False
                vStencil = vApp.Documents.OpenEx("Basic Flowchart Shapes.vss", 4)
                vConnectorStencil = vApp.Documents.OpenEx("Connectors.vss", 4)
                vFlowChartMaster = vStencil.Masters("Process")
                vConnectorMaster = vConnectorStencil.Masters("Side to Side 1")
            End If
            is_VFileSave = False
            PageCounter += 1
            GroupCounter += 1
            Call CreateTblDuplicateCheck(TblCheckDuplication)
            If PageCounter > 1 Then
                vDoc.Pages.Add()
            End If
            'If PageCounter > 1 Then
            Try
                newXlWorkSheet.Columns.AutoFit()
            Catch ex As Exception

            End Try
            'End If

            If PageCounter > 3 Then
                newXlWorkbook.Sheets.Add(, newXlWorkSheet)
            End If
            Call TitleOfGroup(PageCounter)
            Call ColumnNames(GroupCounter, PageCounter)
            Call FirstLayer(TopNumber)
            Call LayeredAnagram(LayerName, PreLayerNumber, PreLayerItem)
            ' vTitleShape.Text = "From:  " & InitialDatetime & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & EndDatetime
            vTitleShape.Text = "From:  " & IntialGrpTime & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & EndGrpTime
            lbCheckedaPraty.Text = CInt(lbTotalNoOfaParty.Text) - RecordsCount()
            lbCheckedaPraty.Refresh()
            lbNoOfAnagram.Text = GroupCounter
            Try
                newXlWorkSheet.Columns.AutoFit()
            Catch ex As Exception

            End Try
        End While
        If is_VFileSave = False Then
            vFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (GroupCounter - PageCounter) + 1 & " to " & GroupCounter & ".vsd"
            xlFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (GroupCounter - PageCounter) + 1 & " to " & GroupCounter & ".xlsx"
            Try
                ' Delete the previous version of the file.
                Kill(vFilePath)
                Kill(xlFilePath)
                newXlWorkSheet.Columns.AutoFit()
            Catch
            End Try
            vDoc.SaveAs(vFilePath)

            newXlWorkbook.SaveAs(xlFilePath)
            newXlWorkbook.Close()
            vDoc.Close()
            vApp.Quit()
            ' newXlWorkbook.Close()
            newXlApp.Quit()
            vDoc = Nothing
            vApp = Nothing
        End If
        MsgBox("File(s) have been saved in : " & vDirectoryPath, MsgBoxStyle.Information)
    End Sub
    Sub ColumnNames(ByVal GroupNo As Integer, ByVal SheetNumber As Integer)
        IndexRow = 1
        Try

            newXlWorkSheet = newXlWorkbook.Sheets(SheetNumber)
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
        newXlWorkSheet.Name = "Group" & GroupNo
        newXlWorkSheet.Cells(1, 1) = "a"
        newXlWorkSheet.Range("A:A").NumberFormat = "###"
        newXlWorkSheet.Cells(1, 2) = "CNIC_a"
        newXlWorkSheet.Range("B:B").NumberFormat = "###"
        newXlWorkSheet.Cells(1, 3) = "b"
        newXlWorkSheet.Range("C:C").NumberFormat = "###"
        newXlWorkSheet.Cells(1, 4) = "CNIC_b"
        newXlWorkSheet.Range("D:D").NumberFormat = "###"
        newXlWorkSheet.Cells(1, 5) = "Time"
        newXlWorkSheet.Range("E:E").NumberFormat = "hh:mm:ss AM/PM"
        newXlWorkSheet.Cells(1, 6) = "Date"
        newXlWorkSheet.Cells(1, 7) = "Call Type"
        newXlWorkSheet.Cells(1, 8) = "Duration"
        newXlWorkSheet.Cells(1, 9) = "Cell ID"
        newXlWorkSheet.Cells(1, 10) = "IMEI"
        newXlWorkSheet.Range("J:J").NumberFormat = "###"
        newXlWorkSheet.Cells(1, 11) = "IMSI"
        newXlWorkSheet.Range("K:K").NumberFormat = "###"
        newXlWorkSheet.Cells(1, 12) = "Site"
        newXlWorkSheet.Range("A1:P1").Font.Bold = True


    End Sub
    Sub FirstLayer(ByVal aParty As String)
        Is_aPartyFound = False
        Dim cmdBuildAnagram As SqlCommand
        Dim readerBuildAnagram As SqlDataReader
        'Dim cmdOuterLayer As SqlCommand
        Dim QueryBuildAnagram As String = Nothing
        Dim bParty As String = Nothing
        Dim QueryOuterLayer As String = Nothing
        Call ConnectionOpen()
        QueryBuildAnagram = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & aParty & "' AND [Call Type] like 'Out%'),(Select max(Time) as EndTime from [tblAnagram] where a like '%" & aParty & "'),(Select min(Time) as IntialTime from [tblAnagram] where a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        readerBuildAnagram = cmdBuildAnagram.ExecuteReader
        While readerBuildAnagram.Read
            Is_aPartyFound = True
            GprCNIC = Nothing
            If aParty.Length >= 10 Or aParty <> "NullValue" Then
                IsCNIC_Need = True
                GprCNIC = GetCNIC_FromExcel(aParty, "CNIC_a", "a")
                'Call FindPhoneNumber(bParty)
                IsCNIC_Need = False
            End If
            'Call FindPhoneNumber(aParty)
            If GprCNIC = Nothing Then
                GprCNIC = "Not Found"
            End If
            ShapeText = aParty & "(" & readerBuildAnagram(0) & "a)" & vbCrLf & GprCNIC & vbCrLf & "In(" & readerBuildAnagram(1) & ") Out(" & readerBuildAnagram(2) & ")"
            'Call InsertGrpLayer(Layer, aParty, aParty, readerBuildAnagram(0), readerBuildAnagram(1), readerBuildAnagram(2))
            'Call addColToAnagrambls("Layer" & AnagramTable.Columns.Count)
            'If SecondTurn = False Then
            IntialGrpTime = readerBuildAnagram(4)
            EndGrpTime = readerBuildAnagram(3)
            'Else
            '    CallTimeResult = DateTime.Compare(IntialGrpTime, readerBuildAnagram(4))
            '    If CallTimeResult = 1 Then
            '        IntialGrpTime = readerBuildAnagram(4)
            '    End If
            '    CallTimeResult = DateTime.Compare(EndGrpTime, readerBuildAnagram(3))
            '    If CallTimeResult = -1 Then
            '        EndGrpTime = readerBuildAnagram(3)
            '    End If
            'End If
            Exit While
        End While
        readerBuildAnagram.Close()
        cmdBuildAnagram.Dispose()
        TargetConnection.Close()
        If Is_aPartyFound = True Then
            Call DropFirstLayer()
        End If
    End Sub
    Dim LayerX As Double = 0.51
    Dim LayerY As Double = 12
    Dim PreLayerNumber As Integer = 0
    Dim PreLayerItem As Integer = 0
    Dim LayerName As String = "L"
    Dim NextLayerItems As Integer = 0
    Dim NexLayerNumber As Integer = 0
    Dim FromShape As Visio.Shape
    Sub DropFirstLayer()
        LayerX = 0.83
        LayerY = 12.0
        NexLayerNumber = 0
        NextLayerItems = 0
        NexLayerNumber += 1
        NextLayerItems += 1
        FromShape = vApp.ActivePage.Drop(vFlowChartMaster, LayerX, LayerY)
        FromShape.Characters.CharProps(2) = &H1
        FromShape.CellsU("Fillforegnd").Formula = "RGB(238, 221, 130)"
        FromShape.Cells("Width").ResultIU = 1.13
        FromShape.Cells("Height").ResultIU = 0.45
        FromShape.Text = ShapeText
        FromShape.Name = ShapeText.Substring(0, ShapeText.IndexOf("(")) & LayerName & NexLayerNumber & NextLayerItems
        PreLayerNumber = NexLayerNumber
        PreLayerItem = NextLayerItems
        Y_lastItem = LayerY
        CurrentItem = NextLayerItems
    End Sub
    Dim AllShapes As Visio.Shapes
    Dim CurrentItem As Integer
    Dim NextItem As Integer = 1
    Dim IsNextShape As Boolean = False
    Dim subNextlayerItem As Integer = 0
    Sub LayeredAnagram(ByVal LayerName As String, ByVal LayerNumber As Integer, ByVal LayerItem As Integer)
        AllShapes = vApp.ActivePage.Shapes
        PreLayerNumber = NexLayerNumber
        PreLayerItem = NextLayerItems
        NextLayerItems = 0
        Y_lastItem = 12.0

        isNewLayerAdd = False
        For j As Integer = 1 To LayerItem
            CurrentItem = j
            For i As Integer = 1 To AllShapes.Count
                If AllShapes.Item(i).Name Like "*" & LayerName & LayerNumber & j Then
                    FromShape = AllShapes.Item(i)
                    IsNextShape = True
                    subNextlayerItem = 0
                    Call AnagramCreation(FromShape.Text.Substring(0, FromShape.Text.IndexOf("(")))
                End If
            Next
        Next
        If isNewLayerAdd = True Then
            Call LayeredAnagram(LayerName, NexLayerNumber, NextLayerItems)
        End If
    End Sub
    Dim Y_lastItem As Double
    Dim aPartyColor As String = "RGB(238, 221, 130)"
    Dim bPartyColor As String = "RGB(255, 228, 225)"
    Sub ToLayerItem(ByVal vShapeText As String, ByVal ShapeColor As String)
        'If RightX > FromShape.Cells("pinx").ResultIU Then
        '    RightY -= 0.48
        'End If
        If IsNextShape = True Then
            If Y_lastItem < FromShape.Cells("piny").ResultIU Then
                NextItem = 1
                For k As Integer = CurrentItem To PreLayerItem
                    For d As Integer = 1 To AllShapes.Count
                        If AllShapes.Item(d).Name Like "*" & LayerName & PreLayerNumber & k Then
                            AllShapes.Item(d).Cells("piny").ResultIU = Y_lastItem - (NextItem * 0.48)
                            NextItem += 1
                        End If
                    Next
                Next

            End If
            IsNextShape = False

        End If
        NextLayerItems += 1
        subNextlayerItem += 1
        Y_lastItem = FromShape.Cells("piny").ResultIU - (subNextlayerItem - 1) * 0.48
        'RightX = FromShape.Cells("pinx").ResultIU + 1.175
        ToShape = vApp.ActivePage.Drop(vFlowChartMaster, FromShape.Cells("pinx").ResultIU + 1.373, FromShape.Cells("piny").ResultIU - (subNextlayerItem - 1) * 0.48)
        ToShape.Characters.CharProps(2) = &H1
        ToShape.CellsU("Fillforegnd").Formula = ShapeColor
        ToShape.Cells("Width").ResultIU = 1.21
        ToShape.Cells("Height").ResultIU = 0.45
        ToShape.Text = vShapeText
        ToShape.Name = vShapeText.Substring(0, vShapeText.IndexOf("(")) & LayerName & NexLayerNumber & NextLayerItems
        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(FromShape.Cells("AlignRight"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(ToShape.Cells("AlignLeft"))
        ' FromShape = ToShape
    End Sub
    Dim ConnOuterLayer As SqlConnection
    'Is_aPartyFound = False
    '    isNewLayerAdd = False
    Dim Conn_b_in_a As SqlConnection
    Dim cmd_b_in_a As SqlCommand
    Dim Reader_b_in_a As SqlDataReader
    Dim Query_b_in_a As String = Nothing
    Dim Is_b_in_a As Boolean = False
    Dim Conn_aTo_b_Del As SqlConnection
    Dim cmd_aTo_b_Del As SqlCommand
    Dim Query_Del_aTo_b As String
    Dim ConnToExcel As SqlConnection
    Dim CmdToExcel As SqlCommand
    Dim ReaderToExcel As SqlDataReader
    Dim QueryToExcel As String
    Dim rowNumber As Integer = Nothing
    Dim CallTimeResult As Integer = Nothing
    Dim cmdBuildAnagram As SqlCommand
    Dim readerBuildAnagram As SqlDataReader
    Dim cmdOuterLayer As SqlCommand
    Dim readerOuterLayer As SqlDataReader
    Dim QueryBuildAnagram As String = Nothing
    Dim bParty As String = Nothing
    Dim QueryOuterLayer As String = Nothing
    Dim Previous_bParty As String
    Dim IndexRow As Long
    Dim connGetCNIC_FromExcel As SqlConnection
    Dim cmdGetCNIC_FromExcel As SqlCommand
    Dim readerGetCNIC_FromEXcel As SqlDataReader
    Dim QueryGetCNIC_FromExcel As String
    Function GetCNIC_FromExcel(ByVal ph_Number As String, ByVal ColCNIC As String, ByVal ColParty As String) As String
        Dim CNIC_FromExcel As String = Nothing
        connGetCNIC_FromExcel = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        connGetCNIC_FromExcel.Open()
        QueryGetCNIC_FromExcel = "Select " & ColCNIC & " from [tblAnagram] where " & ColParty & " Like '%" & ph_Number & "'"
        cmdGetCNIC_FromExcel = New SqlCommand(QueryGetCNIC_FromExcel, connGetCNIC_FromExcel)
        Try
            readerGetCNIC_FromEXcel = cmdGetCNIC_FromExcel.ExecuteReader
            While readerGetCNIC_FromEXcel.Read
                If IsDBNull(readerGetCNIC_FromEXcel(0)) = False Then
                    CNIC_FromExcel = Trim(readerGetCNIC_FromEXcel(0))
                End If
            End While
            readerGetCNIC_FromEXcel.Close()
            cmdGetCNIC_FromExcel.Dispose()
            connGetCNIC_FromExcel.Close()
            connGetCNIC_FromExcel.Dispose()
        Catch ex As Exception
            Try
                readerGetCNIC_FromEXcel.Close()
                cmdGetCNIC_FromExcel.Dispose()
                connGetCNIC_FromExcel.Close()
                connGetCNIC_FromExcel.Dispose()
            Catch ex1 As Exception

            End Try

        End Try
        Return CNIC_FromExcel
    End Function
    Sub AnagramCreation(ByVal aParty As String)

        Previous_bParty = Nothing
        ConnOuterLayer = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        'If ConnOuterLayer.State = ConnectionState.Closed Then
        '    ConnOuterLayer.Open()
        'End If
        Conn_b_in_a = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        'If Conn_b_in_a.State = ConnectionState.Closed Then
        '    Conn_b_in_a.Open()
        'End If
        Conn_aTo_b_Del = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        ConnToExcel = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        'If Conn_aTo_b_Del.State = ConnectionState.Closed Then
        '    Conn_aTo_b_Del.Open()
        'End If
        'Is_aPartyFound = False
        'isNewLayerAdd = False
        'Dim rowNumber As Integer = Nothing
        'Dim CallTimeResult As Integer = Nothing
        'Dim cmdBuildAnagram As SqlCommand
        'Dim readerBuildAnagram As SqlDataReader
        'Dim cmdOuterLayer As SqlCommand
        'Dim readerOuterLayer As SqlDataReader
        'Dim QueryBuildAnagram As String = Nothing
        'Dim bParty As String = Nothing
        'Dim QueryOuterLayer As String = Nothing
        Dim a As String = Nothing
        Dim b As String = Nothing
        Dim grpTime As String = Nothing
        Dim grpDate As String = Nothing
        Dim Duration As String = Nothing

        If aParty = "NullValue" Or aParty.Length < 1 Then 'Or aParty.Length < 10 Then
            ConnOuterLayer.Close()
            Conn_b_in_a.Close()
            Conn_aTo_b_Del.Close()
            Exit Sub
        End If
        ConnToExcel.Open()
        'If aParty = "923067902479" Then
        '    MsgBox("prompt number", MsgBoxStyle.Information)
        'End If
        QueryToExcel = "Select * from [tblAnagram] where a Like '%" & aParty & "'"
        CmdToExcel = New SqlCommand(QueryToExcel, ConnToExcel)
        ReaderToExcel = CmdToExcel.ExecuteReader
        'Inserting data in excel file
        While ReaderToExcel.Read
            If IsDBNull(ReaderToExcel("a")) = False Then
                a = ReaderToExcel("a")
            Else
                a = "NULL"
            End If
            If IsDBNull(ReaderToExcel("b")) = False Then
                b = ReaderToExcel("b")
            Else
                b = "NULL"
            End If
            If IsDBNull(ReaderToExcel("Time")) = False Then
                grpTime = ReaderToExcel("Time")
            Else
                grpTime = "NULL"
            End If
            If IsDBNull(ReaderToExcel("Date")) = False Then
                grpDate = ReaderToExcel("Date")
            Else
                grpDate = "NULL"
            End If
            If IsDBNull(ReaderToExcel("Duration")) = False Then
                Duration = ReaderToExcel("Duration")
            Else
                Duration = "NULL"
            End If

            If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                IndexRow = IndexRow + 1
                For j As Integer = 1 To 12
                    newXlWorkSheet.Cells(IndexRow, j) = ReaderToExcel(j - 1)
                Next
            End If
        End While
        ReaderToExcel.Close()
        CmdToExcel.Dispose()
        ConnToExcel.Close()
        Call ConnectionOpen()
        QueryBuildAnagram = "Select Distinct(b) from [tblAnagram] where a Like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        readerBuildAnagram = cmdBuildAnagram.ExecuteReader
        isTopRow = True

        'rowNumber = GetIndex("Layer" & AnagramTable.Columns.Count - 1, aParty)
        While readerBuildAnagram.Read
            If IsDBNull(readerBuildAnagram(0)) = False Then
                bParty = readerBuildAnagram(0)
                If bParty.Length >= 10 Then
                    bParty = bParty.Substring(bParty.Length - 10, 10)
                End If
            Else
                bParty = ""
            End If
            If bParty = "NullValue" Then 'Or bParty.Length < 10 Then
                QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "'"
            Else
                QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] = 'In'),(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] = 'Out'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "'"
                'QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where  b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' ), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' ) from [tblAnagram] where b like '%" & bParty & "'"
            End If
            'QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "'"
            'QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where  b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' ), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' ) from [tblAnagram] where b like '%" & bParty & "'"
            ConnOuterLayer.Open()
            cmdOuterLayer = New SqlCommand(QueryOuterLayer, ConnOuterLayer)
            readerOuterLayer = cmdOuterLayer.ExecuteReader
            While readerOuterLayer.Read
                Is_b_in_a = False
                If isNewLayerAdd = False Then
                    NexLayerNumber += 1
                    'Call addColToAnagrambls("Layer" & AnagramTable.Columns.Count)
                    isNewLayerAdd = True
                End If
                'ShapeText = bParty & "(" & readerOuterLayer(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & readerOuterLayer(1) & ") Out(" & readerOuterLayer(2) & ")"
                CallTimeResult = DateTime.Compare(IntialGrpTime, readerOuterLayer(4))
                If CallTimeResult = 1 Then
                    IntialGrpTime = readerOuterLayer(4)
                End If
                CallTimeResult = DateTime.Compare(EndGrpTime, readerOuterLayer(3))
                If CallTimeResult = -1 Then
                    EndGrpTime = readerOuterLayer(3)
                End If
                GprCNIC = Nothing
                If bParty.Length >= 10 Or bParty <> "NullValue" Then
                    IsCNIC_Need = True
                    GprCNIC = GetCNIC_FromExcel(bParty, "CNIC_b", "b")
                    'Call FindPhoneNumber(bParty)
                    IsCNIC_Need = False
                End If
                If GprCNIC = Nothing Then
                    GprCNIC = "Not Found"
                End If
                Query_b_in_a = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'Out%') from [tblAnagram] where a like '%" & bParty & "'"
                Conn_b_in_a.Open()
                cmd_b_in_a = New SqlCommand(Query_b_in_a, Conn_b_in_a)
                Reader_b_in_a = cmd_b_in_a.ExecuteReader
                While Reader_b_in_a.Read
                    ShapeText = bParty & "(" & readerOuterLayer(0) & "b," & Reader_b_in_a(0) & "a)" & vbCrLf & GprCNIC & vbCrLf & "In(" & readerOuterLayer(1) + Reader_b_in_a(1) & ") Out(" & readerOuterLayer(2) + Reader_b_in_a(2) & ")"
                    Is_b_in_a = True
                    Exit While
                End While
                Reader_b_in_a.Close()
                cmd_b_in_a.Dispose()
                Conn_b_in_a.Close()
                If Is_b_in_a = False Then
                    ShapeText = bParty & "(" & readerOuterLayer(0) & "b)" & vbCrLf & GprCNIC & vbCrLf & "In(" & readerOuterLayer(1) & ") Out(" & readerOuterLayer(2) & ")"
                End If
                If Is_b_in_a = False Then
                    Call ToLayerItem(ShapeText, bPartyColor)
                Else
                    Call ToLayerItem(ShapeText, aPartyColor)
                End If
                Exit While
            End While
            readerOuterLayer.Close()
            cmdOuterLayer.Dispose()
            ConnOuterLayer.Close()
            Query_Del_aTo_b = "Delete from [tblAnagram] where a Like '%" & aParty & "' AND b like '%" & bParty & "'"
            Conn_aTo_b_Del.Open()
            cmd_aTo_b_Del = New SqlCommand(Query_Del_aTo_b, Conn_aTo_b_Del)
            cmd_aTo_b_Del.ExecuteNonQuery()
            cmd_aTo_b_Del.Dispose()
            Conn_aTo_b_Del.Close()
            Previous_bParty = bParty
        End While
        readerBuildAnagram.Close()
        cmdBuildAnagram.Dispose()
        TargetConnection.Close()
        TargetConnection.Dispose()
        Call ConnectionOpen()
        QueryBuildAnagram = "Delete from [tblAnagram] where a Like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        cmdBuildAnagram.ExecuteNonQuery()
        cmdBuildAnagram.Dispose()
        If aParty <> "NullValue" Or aParty.Length >= 10 Then
            ConnToExcel.Open()
            QueryToExcel = "Select * from [tblAnagram] where b Like '%" & aParty & "'"
            CmdToExcel = New SqlCommand(QueryToExcel, ConnToExcel)
            ReaderToExcel = CmdToExcel.ExecuteReader
            'inserting data in excel file
            While ReaderToExcel.Read
                If IsDBNull(ReaderToExcel("a")) = False Then
                    a = ReaderToExcel("a")
                Else
                    a = "NULL"
                End If
                If IsDBNull(ReaderToExcel("b")) = False Then
                    b = ReaderToExcel("b")
                Else
                    b = "NULL"
                End If
                If IsDBNull(ReaderToExcel("Time")) = False Then
                    grpTime = ReaderToExcel("Time")
                Else
                    grpTime = "NULL"
                End If
                If IsDBNull(ReaderToExcel("Date")) = False Then
                    grpDate = ReaderToExcel("Date")
                Else
                    grpDate = "NULL"
                End If
                If IsDBNull(ReaderToExcel("Duration")) = False Then
                    Duration = ReaderToExcel("Duration")
                Else
                    Duration = "NULL"
                End If

                If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                    IndexRow = IndexRow + 1
                    For j As Integer = 1 To 12
                        newXlWorkSheet.Cells(IndexRow, j) = ReaderToExcel(j - 1)
                    Next
                End If
            End While
            ReaderToExcel.Close()
            CmdToExcel.Dispose()
            ConnToExcel.Close()
            'Call ConnectionOpen()
            QueryBuildAnagram = "Select distinct(a) from [tblAnagram] where b Like '%" & aParty & "'"
            cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
            readerBuildAnagram = cmdBuildAnagram.ExecuteReader
            While readerBuildAnagram.Read
                If IsDBNull(readerBuildAnagram(0)) = False Then
                    bParty = readerBuildAnagram(0)
                    If bParty.Length >= 10 Then
                        bParty = bParty.Substring(bParty.Length - 10, 10)
                        ''If bParty = Previous_bParty Then
                        ''    GoTo NextItration
                        ''End If
                    End If
                Else
                    bParty = "NullValue"
                End If
                If bParty.Length < 10 Or bParty = "NullValue" Then
                    QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "'"
                Else
                    'QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "') from [tblAnagram] where a like '%" & bParty & "'"
                    QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "'"
                End If
                'QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "'"
                'QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "') from [tblAnagram] where a like '%" & bParty & "'"
                ConnOuterLayer.Open()
                cmdOuterLayer = New SqlCommand(QueryOuterLayer, ConnOuterLayer)
                readerOuterLayer = cmdOuterLayer.ExecuteReader
                While readerOuterLayer.Read
                    GprCNIC = Nothing
                    If bParty.Length >= 10 Or bParty <> "NullValue" Then
                        IsCNIC_Need = True
                        GprCNIC = GetCNIC_FromExcel(bParty, "CNIC_a", "a")
                        'Call FindPhoneNumber(bParty)
                        IsCNIC_Need = False
                    End If
                    If GprCNIC = Nothing Then
                        GprCNIC = "Not Found"
                    End If
                    ShapeText = bParty & "(" & readerOuterLayer(0) & "a)" & vbCrLf & GprCNIC & vbCrLf & "In(" & readerOuterLayer(1) & ") Out(" & readerOuterLayer(2) & ")"
                    'Call InsertGrpLayer(Layer, aParty, bParty, readerOuterLayer(0), readerOuterLayer(1), readerOuterLayer(2))
                    CallTimeResult = DateTime.Compare(IntialGrpTime, readerOuterLayer(4))
                    If CallTimeResult = 1 Then
                        IntialGrpTime = readerOuterLayer(4)
                    End If
                    CallTimeResult = DateTime.Compare(EndGrpTime, readerOuterLayer(3))
                    If CallTimeResult = -1 Then
                        EndGrpTime = readerOuterLayer(3)
                    End If
                    ''If isTopRow = True Then
                    ''    Call UpdateRow("Layer" & AnagramTable.Columns.Count - 1, rowNumber, ShapeText)
                    ''Else
                    ''    Call InsertToAnagramtbl("Layer" & AnagramTable.Columns.Count - 1, rowNumber, ShapeText)
                    ''End If
                    ''rowNumber += 1
                    ''isTopRow = False
                    Call ToLayerItem(ShapeText, aPartyColor)
                    If isNewLayerAdd = False Then
                        NexLayerNumber += 1
                        'Call addColToAnagrambls("Layer" & AnagramTable.Columns.Count)
                        isNewLayerAdd = True
                    End If
                    Exit While
                End While

                readerOuterLayer.Close()
                cmdOuterLayer.Dispose()
                ConnOuterLayer.Close()
NextItration:
            End While

            TargetConnection.Close()
            TargetConnection.Dispose()
            cmdBuildAnagram.Dispose()
        End If
    End Sub
    Sub AnagramCreationOnlyExcel(ByVal aParty As String)

        Previous_bParty = Nothing
        Conn_aTo_b_Del = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        ConnToExcel = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        Dim a As String = Nothing
        Dim b As String = Nothing
        Dim grpTime As String = Nothing
        Dim grpDate As String = Nothing
        Dim Duration As String = Nothing

        If aParty = "NullValue" Or aParty.Length < 10 Then
            Conn_aTo_b_Del.Close()
            Exit Sub
        End If
        ConnToExcel.Open()
        If aParty.Length >= 10 Then
            aParty = aParty.Substring(aParty.Length - 10, 10)
        End If
        QueryToExcel = "Select * from [tblAnagram] where a Like '%" & aParty & "'"
        CmdToExcel = New SqlCommand(QueryToExcel, ConnToExcel)
        ReaderToExcel = CmdToExcel.ExecuteReader
        'Inserting data in excel file
        While ReaderToExcel.Read
            If IsDBNull(ReaderToExcel("a")) = False Then
                a = ReaderToExcel("a")
            Else
                a = "NULL"
            End If
            If IsDBNull(ReaderToExcel("b")) = False Then
                b = ReaderToExcel("b")
                If Trim(b).Length >= 10 Then
                    If lstPhoneNumbers.Contains(Trim(b).Substring(Trim(b).Length - 10, 10)) = False Then
                        lstPhoneNumbers.Add(Trim(b).Substring(Trim(b).Length - 10, 10))
                    End If
                End If
            Else
                b = "NULL"
            End If
            If IsDBNull(ReaderToExcel("Time")) = False Then
                grpTime = ReaderToExcel("Time")
            Else
                grpTime = "NULL"
            End If
            If IsDBNull(ReaderToExcel("Date")) = False Then
                grpDate = ReaderToExcel("Date")
            Else
                grpDate = "NULL"
            End If
            If IsDBNull(ReaderToExcel("Duration")) = False Then
                Duration = ReaderToExcel("Duration")
            Else
                Duration = "NULL"
            End If

            If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                IndexRow = IndexRow + 1
                For j As Integer = 1 To 12
                    newXlWorkSheet.Cells(IndexRow, j) = ReaderToExcel(j - 1)
                Next
            End If
        End While
        ReaderToExcel.Close()
        CmdToExcel.Dispose()
        ConnToExcel.Close()

        Call ConnectionOpen()
        QueryBuildAnagram = "Delete from [tblAnagram] where a Like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        cmdBuildAnagram.ExecuteNonQuery()
        cmdBuildAnagram.Dispose()
        TargetConnection.Close()
        If aParty <> "NullValue" Or aParty.Length >= 1 Then
            ConnToExcel.Open()
            QueryToExcel = "Select * from [tblAnagram] where b Like '%" & aParty & "'"
            CmdToExcel = New SqlCommand(QueryToExcel, ConnToExcel)
            ReaderToExcel = CmdToExcel.ExecuteReader
            'inserting data in excel file
            While ReaderToExcel.Read
                If IsDBNull(ReaderToExcel("a")) = False Then
                    a = ReaderToExcel("a")

                    If Trim(a).Length >= 10 Then
                        If lstPhoneNumbers.Contains(Trim(a).Substring(Trim(a).Length - 10, 10)) = False Then
                            lstPhoneNumbers.Add(Trim(a).Substring(Trim(a).Length - 10, 10))
                        End If
                    End If
                Else
                    a = "NULL"
                End If
                If IsDBNull(ReaderToExcel("b")) = False Then
                    b = ReaderToExcel("b")

                Else
                    b = "NULL"
                End If
                If IsDBNull(ReaderToExcel("Time")) = False Then
                    grpTime = ReaderToExcel("Time")
                Else
                    grpTime = "NULL"
                End If
                If IsDBNull(ReaderToExcel("Date")) = False Then
                    grpDate = ReaderToExcel("Date")
                Else
                    grpDate = "NULL"
                End If
                If IsDBNull(ReaderToExcel("Duration")) = False Then
                    Duration = ReaderToExcel("Duration")
                Else
                    Duration = "NULL"
                End If

                If IsDuplicate(TblCheckDuplication, a, b, grpTime, grpDate, Duration) = False Then
                    IndexRow = IndexRow + 1
                    For j As Integer = 1 To 12
                        newXlWorkSheet.Cells(IndexRow, j) = ReaderToExcel(j - 1)
                    Next
                End If

            End While

            ReaderToExcel.Close()
            CmdToExcel.Dispose()
            ConnToExcel.Close()

            Call ConnectionOpen()
            QueryBuildAnagram = "Delete from [tblAnagram] where b Like '%" & aParty & "'"
            cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
            cmdBuildAnagram.ExecuteNonQuery()
            cmdBuildAnagram.Dispose()
            TargetConnection.Close()
            Call ConnectionOpen()
        End If
        TargetConnection.Close()
        TargetConnection.Dispose()
        cmdBuildAnagram.Dispose()
    End Sub
    Sub LayeringAnagramXL()
        Dim PageCounter As Integer = 0
        
        Dim vDirectoryPath As String = System.IO.Path.GetDirectoryName(lbGroupSheetPath.Text)
        Dim vFilePath As String = Nothing
        Dim is_VFileSave As Boolean = False
        Dim vFileCounter As Integer = 1
        Dim xlFilePath As String = Nothing
        GroupCounter = 0
        lbTotalNoOfaParty.Text = "Initializing..."
        lbCheckedaPraty.Text = "Initializing..."
        lbNoOfAnagram.Text = "Initializing.."
        Try
            Call CreateAnagramTable("tblAnagram")
        Catch ex As Exception
            MsgBox("Error in Creating Anagram Table", MsgBoxStyle.Information)
            Exit Sub
        End Try
        Try
            Call BTS_to_SQL()
        Catch ex As Exception
            MsgBox("Error in BTS to SQL" & vbCrLf & Err.Description, MsgBoxStyle.Information)
            Exit Sub
        End Try

        lbTotalNoOfaParty.Text = RecordsCount()
        ' Dim vDirectoryPath As String = System.IO.Path.GetDirectoryName(lbGroupSheetPath.Text)
        If RecordsCount() > 0 Then
            Try
                'creating excel file
                Call CreateAnagramExcelFile()
            Catch ex As Exception
                MsgBox("Error in Creating Excel file", MsgBoxStyle.Information)
                Exit Sub
            End Try
        End If

        While RecordsCount() >= 1
            If PageCounter = 25 Then

                xlFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (PageCounter * (vFileCounter - 1)) + 1 & " to " & GroupCounter & ".xlsx"
                Try
                    Kill(xlFilePath)
                    newXlWorkSheet.Columns.AutoFit()
                Catch
                End Try
                vFileCounter += 1

                newXlWorkbook.SaveAs(xlFilePath)
                newXlWorkbook.Close()
                is_VFileSave = True
                PageCounter = 0
                newXlWorkbook = newXlApp.Workbooks.Add()
            End If
            is_VFileSave = False
            PageCounter += 1
            GroupCounter += 1
            Call CreateTblDuplicateCheck(TblCheckDuplication)

            'If PageCounter > 1 Then
            Try
                newXlWorkSheet.Columns.AutoFit()
            Catch ex As Exception

            End Try
            'End If

            If PageCounter > 3 Then
                newXlWorkbook.Sheets.Add(, newXlWorkSheet)
            End If

            Call ColumnNames(GroupCounter, PageCounter)
            'Call FirstLayer(TopNumber)
            Call LayeredAnagramXL(TopNumberXL)
            ' vTitleShape.Text = "From:  " & InitialDatetime & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & EndDatetime
            lbCheckedaPraty.Text = CInt(lbTotalNoOfaParty.Text) - RecordsCount()
            lbNoOfAnagram.Text = GroupCounter
            Try
                newXlWorkSheet.Columns.AutoFit()
            Catch ex As Exception

            End Try
        End While
        If is_VFileSave = False Then
            xlFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (GroupCounter - PageCounter) + 1 & " to " & GroupCounter & ".xlsx"
            Try
                ' Delete the previous version of the file.

                Kill(xlFilePath)
                newXlWorkSheet.Columns.AutoFit()
            Catch
            End Try
            newXlWorkbook.SaveAs(xlFilePath)
            newXlWorkbook.Close()
            ' newXlWorkbook.Close()
            newXlApp.Quit()

        End If
        MsgBox("File(s) have been saved in : " & vDirectoryPath, MsgBoxStyle.Information)
    End Sub
    Dim lstPhoneNumbers As New List(Of String)()
    Sub LayeredAnagramXL(Optional aParty As String = Nothing)
        'lstPhoneNumbers.Add(aParty)

        For i As Integer = 0 To lstPhoneNumbers.Count
            If lstPhoneNumbers.Count >= 1 Then
                aParty = lstPhoneNumbers.Item(0)
                lstPhoneNumbers.RemoveAt(0)
                Call AnagramCreationOnlyExcel(aParty)
                'lstPhoneNumbers.Remove(aParty)
            End If
        Next
        If lstPhoneNumbers.Count >= 1 Then
            Call LayeredAnagramXL()
        End If

    End Sub
    Function TopNumberXL() As String
        Dim aParty As String = Nothing
        Dim cmdBuildAnagram As SqlCommand
        Dim readerBuildAnagram As SqlDataReader
        Dim QueryBuildAnagram As String = "Select TOP 1 [a] from tblAnagram"
        Call ConnectionOpen()
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        readerBuildAnagram = cmdBuildAnagram.ExecuteReader
        While readerBuildAnagram.Read
            If IsDBNull(readerBuildAnagram(0)) = False Then
                aParty = readerBuildAnagram(0)
                If aParty.Length >= 10 Then
                    aParty = aParty.Substring(aParty.Length - 10, 10)
                End If
            End If
            Exit While
        End While
        readerBuildAnagram.Close()
        cmdBuildAnagram.Dispose()
        TargetConnection.Close()
        TargetConnection.Dispose()
        lstPhoneNumbers.Add(aParty)
        Return aParty
    End Function
    Private Sub btnBuildAnagram_Click(sender As System.Object, e As System.EventArgs) Handles btnBuildAnagram.Click
        If lbGroupSheetPath.Text.Trim = "" Then
            MsgBox("Please select the file to proccess")
            Exit Sub
        End If
        If cbkWithVisio.Checked = False Then
            Call LayeringAnagramXL()
        Else
            Call LayeringAnagram()
        End If
        'Call CreateAnagramTable("tblAnagram")
        'Call BTS_to_SQL()
        'Call CreatetblAnagram(tracingTable)
        ''Call createGrpLayer(Layer_Name & layerCounter)
        'If RecordsCount() > 0 Then
        '    Call CreateVisioDoc()
        'End If
        'Dim PageCounter As Integer = 0
        'Dim vDirectoryPath As String = System.IO.Path.GetDirectoryName(lbGroupSheetPath.Text)
        'Dim vFilePath As String = Nothing
        'Dim is_VFileSave As Boolean = False
        'Dim vFileCounter As Integer = 1
        'GroupCounter = 0
        'While RecordsCount() >= 1
        '    'Call addColToAnagrambls("Layer1")
        '    Call DeleteTblCols()
        '    Call BuildAnagram(TopNumber, False)
        '    Call RecursiveSearch()
        '    If AnagramTable.Rows.Count > 0 Then
        '        If PageCounter = 25 Then
        '            vFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (PageCounter * (vFileCounter - 1)) + 1 & " to " & GroupCounter & ".vsd"
        '            Try
        '                ' Delete the previous version of the file.
        '                Kill(vFilePath)
        '            Catch
        '            End Try
        '            vFileCounter += 1
        '            vDoc.SaveAs(vFilePath)
        '            is_VFileSave = True
        '            PageCounter = 0
        '            vDoc.Close()
        '            vDoc = vApp.Documents.Add("")
        '            vDoc.PaperSize = Visio.VisPaperSizes.visPaperSizeLegal
        '            'vApp.Visible = False
        '            vStencil = vApp.Documents.OpenEx("Basic Flowchart Shapes.vss", 4)
        '            vConnectorStencil = vApp.Documents.OpenEx("Connectors.vss", 4)
        '            vFlowChartMaster = vStencil.Masters("Process")
        '            vConnectorMaster = vConnectorStencil.Masters("Side to Side 1")
        '        End If
        '        is_VFileSave = False
        '        PageCounter += 1
        '        GroupCounter += 1
        '        If PageCounter > 1 Then
        '            vDoc.Pages.Add()
        '        End If
        '        Call TitleOfGroup(PageCounter)
        '        Call CreateAnagram()

        '    End If

        'End While
        'If is_VFileSave = False Then
        '    vFilePath = vDirectoryPath & "\Anagram of groups on BTS " & (GroupCounter - PageCounter) + 1 & " to " & GroupCounter & ".vsd"
        '    Try
        '        ' Delete the previous version of the file.
        '        Kill(vFilePath)
        '    Catch
        '    End Try
        '    vDoc.SaveAs(vFilePath)
        '    vDoc.Close()
        '    vApp.Quit()
        '    vDoc = Nothing
        '    vApp = Nothing
        'End If

        'vTitleShape.Text = "From:  " & InitialDatetime & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & EndDatetime



        GC.Collect()


    End Sub
    Dim ConnectionLayer As SqlConnection
    Dim cmdLayers As SqlCommand
    Dim readerLayers As SqlDataReader
    Dim LayersQuery As String = Nothing
    Dim isNewLayerAdd As Boolean = False
    Dim TrailName As String = "T"
    Dim TrailNo As Integer = 0
    Dim GroupShapes As Visio.Shapes
    ' Dim FromShape As Visio.Shape
    Dim ToShape As Visio.Shape
    Dim ItemName As String = Nothing
    Dim CentralShapeName As String
    Sub CreateAnagram()
        'Call CreateVisioDoc()
        ' Call TitleOfGroup(1)
        vTitleShape.Text = "From:  " & IntialGrpTime & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & EndGrpTime
        TrailNo = 0
        LeftX = 4.25
        LeftY = 12.0
        RightX = 4.25
        RightY = 12.0
        If AnagramTable.Rows.Count > 0 Then
            ShapeText = AnagramTable.Rows(0)(0)
            CentralShapeName = ShapeText.Substring(0, ShapeText.IndexOf("(")) & "Central"
            Call DropCentralShape()
        End If
        For i As Integer = 0 To AnagramTable.Rows.Count - 1
            GroupShapes = vApp.ActivePage.Shapes
            For j As Integer = 0 To AnagramTable.Columns.Count - 1
                If IsDBNull(AnagramTable.Rows(i)(j)) = False Then
                    ShapeText = AnagramTable.Rows(i)(j)
                    ItemName = ShapeText.Substring(0, ShapeText.IndexOf("("))
                    If j = 0 Then
                        FromShape = GroupShapes.Item(ItemName & "Central")
                    Else
                        Try
                            FromShape = GroupShapes.Item(ItemName & TrailName & TrailNo)
                        Catch ex As Exception
                            If FromShape.Name = CentralShapeName Then
                                TrailNo += 1
                                If CLng(TrailNo) Mod 2 > 0 Then
                                    Call DropLeftTrail()
                                Else
                                    Call DropRightTrail()
                                End If
                            ElseIf CLng(TrailNo) Mod 2 > 0 Then
                                Call DropLeftTrail()
                            Else
                                Call DropRightTrail()
                            End If
                        End Try

                    End If

                    ' Call DropLeftTrail()

                End If
            Next
        Next
    End Sub
    Dim LeftX As Double = 4.25
    Dim LeftY As Double = 12.0
    Dim ShapeToPlace As Visio.Shape
    Sub DropLeftTrail()
        If LeftX < FromShape.Cells("pinx").ResultIU Then
            LeftY -= 0.48
        End If
        LeftX = FromShape.Cells("pinx").ResultIU - 1.175
        ShapeToPlace = vApp.ActivePage.Drop(vFlowChartMaster, LeftX, LeftY)
        ShapeToPlace.Characters.CharProps(2) = &H1
        ShapeToPlace.CellsU("Fillforegnd").Formula = "RGB(255, 228, 225)"
        ShapeToPlace.Cells("Width").ResultIU = 0.95
        ShapeToPlace.Cells("Height").ResultIU = 0.45
        ShapeToPlace.Text = ShapeText
        ShapeToPlace.Name = ShapeText.Substring(0, ShapeText.IndexOf("(")) & TrailName & TrailNo
        Dim mytext As String = ShapeToPlace.Text

        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(FromShape.Cells("AlignLeft"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(ShapeToPlace.Cells("AlignRight"))
        FromShape = ShapeToPlace
    End Sub
    Dim RightX As Double = 4.25
    Dim RightY As Double = 12.0
    Sub DropRightTrail()
        If RightX > FromShape.Cells("pinx").ResultIU Then
            RightY -= 0.48
        End If
        RightX = FromShape.Cells("pinx").ResultIU + 1.175
        ToShape = vApp.ActivePage.Drop(vFlowChartMaster, RightX, RightY)
        ToShape.Characters.CharProps(2) = &H1
        ToShape.CellsU("Fillforegnd").Formula = "RGB(255, 228, 225)"
        ToShape.Cells("Width").ResultIU = 0.95
        ToShape.Cells("Height").ResultIU = 0.45
        ToShape.Text = ShapeText
        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(FromShape.Cells("AlignRight"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(ToShape.Cells("AlignLeft"))
        FromShape = ToShape
        'If EmptyValue = True Then
        ' Y_RS_Shape -= 0.55
        'End If
    End Sub
    Sub ConnectionOpenLayer()
        ConnectionLayer = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If ConnectionLayer.State <> ConnectionState.Open Then
            ConnectionLayer.Open()
        End If
    End Sub
    Function RecordsCount() As Integer
        Dim Count As Long = Nothing
        Call ConnectionOpen()
        LayersQuery = "Select Count(*) from  tblAnagram"
        cmdLayers = New SqlCommand(LayersQuery, TargetConnection)
        Count = cmdLayers.ExecuteScalar
        cmdLayers.Dispose()
        TargetConnection.Close()
        Return Count
    End Function
    Sub b_partyLayers()
        Call ConnectionOpenLayer()
        LayersQuery = "Select Distinct(Child) from " & Layer_Name & layerCounter & ""
        cmdLayers = New SqlCommand(LayersQuery, ConnectionLayer)
        readerLayers = cmdLayers.ExecuteReader
        While readerLayers.Read

        End While
    End Sub
    Function TopNumber() As String
        Dim aParty As String = Nothing
        Dim cmdBuildAnagram As SqlCommand
        Dim readerBuildAnagram As SqlDataReader
        Dim QueryBuildAnagram As String = "Select TOP 1 [a] from tblAnagram"
        Call ConnectionOpen()
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        readerBuildAnagram = cmdBuildAnagram.ExecuteReader
        While readerBuildAnagram.Read
            If IsDBNull(readerBuildAnagram(0)) = False Then
                aParty = readerBuildAnagram(0)
                If aParty.Length >= 10 Then
                    aParty = aParty.Substring(aParty.Length - 10, 10)
                End If
            End If
            Exit While
        End While
        readerBuildAnagram.Close()
        cmdBuildAnagram.Dispose()
        TargetConnection.Close()
        TargetConnection.Dispose()
        Return aParty
    End Function
    Dim Left_bParty1() As String = Nothing
    Dim Left_bParty2() As String = Nothing
    Dim Right_bParty1() As String = Nothing
    Dim Right_bParty2() As String = Nothing
    Dim AnagramTable As DataTable
    Dim NewColumn As DataColumn
    Dim ColNumber As Integer = Nothing
    Sub CreatetblAnagram(ByVal tblAnagram As String)
        AnagramTable = New DataTable
        AnagramTable.Columns.Add("Layer0", Type.GetType("System.String"))
        'AnagramTable.Columns.Add("Layer1", Type.GetType("System.String"))
        ''Try
        ''    Call ConnectionOpen()
        ''    QueryString = "IF OBJECT_ID('dbo." & tblAnagram & "') IS NOT NULL DROP TABLE " & tblAnagram & ""
        ''    CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
        ''    CreateTbCommand.ExecuteNonQuery()
        ''    TargetConnection.Close()
        ''Catch ex As Exception
        ''    If TargetConnection.State <> ConnectionState.Closed Then
        ''        TargetConnection.Close()
        ''    End If
        ''End Try
        ''Call ConnectionOpen()
        ''Try
        ''    QueryString = "CREATE TABLE " & tblAnagram & "(Layer01 text)"
        ''    CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
        ''    CreateTbCommand.ExecuteNonQuery()
        ''    TargetConnection.Close()
        ''    TargetConnection.Dispose()
        ''    CreateTbCommand.Dispose()
        ''Catch ex As Exception

        ''End Try

    End Sub
    Sub addColToAnagrambls(ByVal ColName As String)
        NewColumn = New Data.DataColumn(ColName, Type.GetType("System.String"))
        'NewColumn.DefaultValue = ""
        AnagramTable.Columns.Add(NewColumn)
        ''Call ConnectionOpen()
        ''Try
        ''    QueryString = "ALTER TABLE " + tracingTable + " ADD " + ColName + " text"
        ''    CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
        ''    CreateTbCommand.ExecuteNonQuery()
        ''    TargetConnection.Close()
        ''    TargetConnection.Dispose()
        ''    CreateTbCommand.Dispose()
        ''Catch ex As Exception

        ''End Try
    End Sub
    Sub DeleteTblCols()

        Dim countcols As Integer = Nothing
        Try
            AnagramTable.Clear()
            countcols = AnagramTable.Columns.Count
            If countcols > 1 Then
                For i As Integer = countcols To 1 Step -1
                    AnagramTable.Columns.RemoveAt(i)
                Next
            End If
        Catch ex As Exception

        End Try


    End Sub
    Dim RowsCount As Long = Nothing
    Dim ResultRows() As DataRow
    Dim Number As String = Nothing
    Dim ColumnNumber As Integer = Nothing
    Dim CriteriaExpression As String = Nothing
    Sub RecursiveSearch()
        CriteriaExpression = "Layer0 = " & AnagramTable.Rows(0)(0)
        ResultRows = AnagramTable.Select()
        ColumnNumber = AnagramTable.Columns.Count - 1
        For i As Long = 0 To ResultRows.Count - 1
            If IsDBNull(ResultRows(i)(ColumnNumber)) = False Then
                Number = ResultRows(i)(ColumnNumber)
                Number = Number.Substring(0, Number.IndexOf("("))
                Call BuildAnagram(Number, True)
            End If
        Next
        If isNewLayerAdd = True Then
            Call RecursiveSearch()
        End If
    End Sub
    Dim OnlyPhnNumber As String = Nothing
    Function GetIndex(ByVal LayerName As String, ByVal bPartyValue As String) As Integer

        For i As Integer = 0 To AnagramTable.Rows.Count - 1
            If IsDBNull(AnagramTable.Rows(i)(LayerName)) = False Then
                OnlyPhnNumber = AnagramTable.Rows(i)(LayerName)
                If OnlyPhnNumber.Substring(0, 10) = bPartyValue Then
                    Return i
                    Exit For
                End If
            End If
        Next

    End Function

    Dim ConnInsert As SqlConnection
    Dim cmdInsert As SqlCommand
    Dim InsertLayerQuery As String = Nothing
    Sub InsertToAnagramtbl(ByVal LayerName As String, ByVal RowIndex As Integer, ByVal PartyValue As String)
        AddRow = AnagramTable.NewRow
        ColNumber = AnagramTable.Columns.Count
        AddRow(LayerName) = PartyValue
        If ColNumber > 1 Then
            For i As Integer = 0 To ColNumber - 2
                If RowIndex > 0 Then
                    AddRow(i) = AnagramTable.Rows(RowIndex - 1)(i)
                End If
            Next
        End If
        AnagramTable.Rows.InsertAt(AddRow, RowIndex)
        'ConnInsert = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        'If ConnInsert.State <> ConnectionState.Open Then
        '    ConnInsert.Open()
        'End If
        'InsertLayerQuery = "INSERT INTO [" & tracingTable & "] VALUES ('" & a & "','" & b & "','" & Time & "','" & CallType & "')"
        'InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        'InsertIMEICommand.ExecuteNonQuery()
        'TargetConnection.Close()
        'TargetConnection.Dispose()
        'InsertIMEICommand.Dispose()
    End Sub
    Sub UpdateRow(ByVal LayerName As String, ByVal RowIndex As Integer, ByVal bPartyValue As String)
        AnagramTable.Rows(RowIndex).Item(LayerName) = bPartyValue
    End Sub
    Dim isTopRow As Boolean = True
    Dim Is_aPartyFound As Boolean = False
    Dim IntialGrpTime As Date = Nothing
    Dim EndGrpTime As Date = Nothing
    Sub BuildAnagram(ByVal aParty As String, ByVal SecondTurn As Boolean)
        Dim ConnOuterLayer As SqlConnection
        ConnOuterLayer = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        If ConnOuterLayer.State <> ConnectionState.Open Then
            ConnOuterLayer.Open()
        End If
        Is_aPartyFound = False
        isNewLayerAdd = False
        Dim rowNumber As Integer = Nothing
        Dim CallTimeResult As Integer = Nothing
        Dim cmdBuildAnagram As SqlCommand
        Dim readerBuildAnagram As SqlDataReader
        Dim cmdOuterLayer As SqlCommand
        Dim readerOuterLayer As SqlDataReader
        Dim QueryBuildAnagram As String = Nothing
        Dim bParty As String = Nothing
        Dim QueryOuterLayer As String = Nothing
        ''Call ConnectionOpen()
        ''QueryBuildAnagram = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & aParty & "' AND [Call Type] like 'Out%'),(Select max(Time) as EndTime from [tblAnagram] where a like '%" & aParty & "'),(Select min(Time) as IntialTime from [tblAnagram] where a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "'"
        ''cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        ''readerBuildAnagram = cmdBuildAnagram.ExecuteReader
        ''While readerBuildAnagram.Read
        ''    Is_aPartyFound = True
        ''    ShapeText = aParty & "(" & readerBuildAnagram(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & readerBuildAnagram(1) & ") Out(" & readerBuildAnagram(2) & ")"
        ''    'Call InsertGrpLayer(Layer, aParty, aParty, readerBuildAnagram(0), readerBuildAnagram(1), readerBuildAnagram(2))
        ''    'Call addColToAnagrambls("Layer" & AnagramTable.Columns.Count)
        ''    If SecondTurn = False Then
        ''        IntialGrpTime = readerBuildAnagram(4)
        ''        EndGrpTime = readerBuildAnagram(3)
        ''    Else
        ''        CallTimeResult = DateTime.Compare(IntialGrpTime, readerBuildAnagram(4))
        ''        If CallTimeResult = 1 Then
        ''            IntialGrpTime = readerBuildAnagram(4)
        ''        End If
        ''        CallTimeResult = DateTime.Compare(EndGrpTime, readerBuildAnagram(3))
        ''        If CallTimeResult = -1 Then
        ''            EndGrpTime = readerBuildAnagram(3)
        ''        End If
        ''    End If
        ''    Exit While
        ''End While
        ''readerBuildAnagram.Close()
        ''cmdBuildAnagram.Dispose()
        ''TargetConnection.Close()
        ''TargetConnection.Dispose()
        ' ''Call DropCentralShape()
        ''If Is_aPartyFound = False Then
        ''    Exit Sub
        ''Else
        ''    If SecondTurn = False Then
        ''        InsertToAnagramtbl("Layer0", 0, ShapeText)
        ''    End If
        ''End If
        Call ConnectionOpen()
        QueryBuildAnagram = "Select Distinct(b) from [tblAnagram] where a Like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        readerBuildAnagram = cmdBuildAnagram.ExecuteReader

        isTopRow = True
        rowNumber = GetIndex("Layer" & AnagramTable.Columns.Count - 1, aParty)
        While readerBuildAnagram.Read
            If IsDBNull(readerBuildAnagram(0)) = False Then
                bParty = readerBuildAnagram(0)
                If bParty.Length >= 10 Then
                    bParty = bParty.Substring(bParty.Length - 10, 10)
                End If
            Else
                bParty = ""
            End If
            If bParty = "NULL VALUE" Or bParty.Length < 10 Then
                QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "'"
            Else
                QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] = 'In'),(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] = 'Out'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "'"
                'QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where  b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' ), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' ) from [tblAnagram] where b like '%" & bParty & "'"
            End If
            'QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' AND a like '%" & aParty & "') from [tblAnagram] where a like '%" & aParty & "' AND b like '%" & bParty & "'"
            'QueryOuterLayer = "Select (Select Count(b) as TotalCalls from [tblAnagram] where b like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where  b like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where b like '%" & bParty & "' ), (Select min(Time) as InitialTime from [tblAnagram] where b like '%" & bParty & "' ) from [tblAnagram] where b like '%" & bParty & "'"
            cmdOuterLayer = New SqlCommand(QueryOuterLayer, ConnOuterLayer)
            readerOuterLayer = cmdOuterLayer.ExecuteReader
            While readerOuterLayer.Read
                If isNewLayerAdd = False Then
                    Call addColToAnagrambls("Layer" & AnagramTable.Columns.Count)
                    isNewLayerAdd = True
                End If
                ShapeText = bParty & "(" & readerOuterLayer(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & readerOuterLayer(1) & ") Out(" & readerOuterLayer(2) & ")"
                CallTimeResult = DateTime.Compare(IntialGrpTime, readerOuterLayer(4))
                If CallTimeResult = 1 Then
                    IntialGrpTime = readerOuterLayer(4)
                End If
                CallTimeResult = DateTime.Compare(EndGrpTime, readerOuterLayer(3))
                If CallTimeResult = -1 Then
                    EndGrpTime = readerOuterLayer(3)
                End If
                'Call InsertGrpLayer(Layer, aParty, bParty, readerOuterLayer(0), readerOuterLayer(1), readerOuterLayer(2))
                If isTopRow = True Then
                    Call UpdateRow("Layer" & AnagramTable.Columns.Count - 1, rowNumber, ShapeText)
                Else
                    Call InsertToAnagramtbl("Layer" & AnagramTable.Columns.Count - 1, rowNumber, ShapeText)
                End If
                rowNumber += 1
                isTopRow = False
                Exit While
            End While
            readerOuterLayer.Close()
            cmdOuterLayer.Dispose()
        End While
        readerBuildAnagram.Close()
        cmdBuildAnagram.Dispose()
        TargetConnection.Close()
        TargetConnection.Dispose()
        Call ConnectionOpen()
        QueryBuildAnagram = "Delete from [tblAnagram] where a Like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        cmdBuildAnagram.ExecuteNonQuery()
        cmdBuildAnagram.Dispose()
        QueryBuildAnagram = "Select distinct(a) from [tblAnagram] where b Like '%" & aParty & "'"
        cmdBuildAnagram = New SqlCommand(QueryBuildAnagram, TargetConnection)
        readerBuildAnagram = cmdBuildAnagram.ExecuteReader
        While readerBuildAnagram.Read
            If IsDBNull(readerBuildAnagram(0)) = False Then
                bParty = readerBuildAnagram(0)
                If bParty.Length >= 10 Then
                    bParty = bParty.Substring(bParty.Length - 10, 10)
                End If
            Else
                bParty = ""
            End If
            If bParty.Length < 10 Or bParty = "NULL VALUE" Then
                QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "'"
            Else
                'QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "') from [tblAnagram] where a like '%" & bParty & "'"
                QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "'"
            End If
            'QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "' AND b like '%" & aParty & "') from [tblAnagram] where b like '%" & aParty & "' AND a like '%" & bParty & "'"
            'QueryOuterLayer = "Select (Select Count(a) as TotalCalls from [tblAnagram] where a like '%" & bParty & "') ,(Select Count([Call Type]) as Incomings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'In%'),(Select Count([Call Type]) as Outgoings from [tblAnagram] where a like '%" & bParty & "' AND [Call Type] like 'Out%'), (Select max(Time) as EndTime from [tblAnagram] where a like '%" & bParty & "'), (Select min(Time) as InitialTime from [tblAnagram] where a like '%" & bParty & "') from [tblAnagram] where a like '%" & bParty & "'"
            cmdOuterLayer = New SqlCommand(QueryOuterLayer, ConnOuterLayer)
            readerOuterLayer = cmdOuterLayer.ExecuteReader
            While readerOuterLayer.Read
                ShapeText = bParty & "(" & readerOuterLayer(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & readerOuterLayer(1) & ") Out(" & readerOuterLayer(2) & ")"
                'Call InsertGrpLayer(Layer, aParty, bParty, readerOuterLayer(0), readerOuterLayer(1), readerOuterLayer(2))
                CallTimeResult = DateTime.Compare(IntialGrpTime, readerOuterLayer(4))
                If CallTimeResult = 1 Then
                    IntialGrpTime = readerOuterLayer(4)
                End If
                CallTimeResult = DateTime.Compare(EndGrpTime, readerOuterLayer(3))
                If CallTimeResult = -1 Then
                    EndGrpTime = readerOuterLayer(3)
                End If
                ''If isTopRow = True Then
                ''    Call UpdateRow("Layer" & AnagramTable.Columns.Count - 1, rowNumber, ShapeText)
                ''Else
                ''    Call InsertToAnagramtbl("Layer" & AnagramTable.Columns.Count - 1, rowNumber, ShapeText)
                ''End If
                ''rowNumber += 1
                ''isTopRow = False
                Exit While
            End While
            readerOuterLayer.Close()
            cmdOuterLayer.Dispose()
        End While

        TargetConnection.Close()
        TargetConnection.Dispose()
        cmdBuildAnagram.Dispose()
    End Sub

    Sub oldAnagramDesign()
        If lbGroupSheetPath.Text = "" Then
            ' Call CreateAnagramTable("AnagramTable")
            MsgBox("Please Select the File to generate anagram", MsgBoxStyle.Information)
            Exit Sub
        End If
        Call CreateAnagramTable("tblAnagram")
        Call BTS_to_SQL()
        Dim vDirectoryPath As String = System.IO.Path.GetDirectoryName(lbGroupSheetPath.Text)
        Dim vFilePath As String = Nothing
        Dim is_VFileSave As Boolean = False
        Dim ConnectionString As String = Nothing
        Dim Conn_bParty As OleDbConnection
        Dim ConnCounter_bParty As OleDbConnection
        Dim cmd_bParty As OleDbCommand
        Dim Reader_bParty As OleDbDataReader
        Dim cmd_Counter_bPary As OleDbCommand
        Dim ReaderCounter_bParty As OleDbDataReader
        Dim QueryCounter_bParty As String = Nothing
        Dim Conn_b_to_a As OleDbConnection
        Dim ConnCount_b_to_a As OleDbConnection
        Dim cmd_b_to_a As OleDbCommand
        Dim Reader_b_to_a As OleDbDataReader
        Dim cmdCount_b_to_a As OleDbCommand
        Dim ReaderCount_b_to_a As OleDbDataReader

        Dim QueryString_b_to_a As String = Nothing
        Dim QueryStringCount_bToa As String = Nothing
        Dim pageCounter As Integer = 1
        Dim vFileCounter As Integer = 0
        Dim TotalNoOf_a_Party As Integer = Nothing
        Dim cmdOnly_a_party As OleDbCommand
        Call CreateAnagramTable()
        Dim aParty As String = Nothing
        Dim bParty As String = Nothing
        Dim bToaParty As String = Nothing
        Dim InitialDatetime As Date
        Dim EndDatetime As Date
        Dim CallDateTime As Date
        Dim CallTimeResult As Integer = Nothing
        Dim TotalRecords As Integer = Nothing
        Dim extractingQuery As String = Nothing
        Dim totalcalls_a As Integer = Nothing
        Dim Incommings_a As Integer = Nothing
        Dim Outgoings_a As Integer = Nothing
        Dim RemaingCalls As Integer = Nothing
        Dim Query_bParty As String = Nothing
        Dim Query_bToaParty As String = Nothing
        Dim bCounter As Integer = 0
        Dim ShapCounter As Integer = 1
        Dim Only_a_prtyReader As OleDbDataReader

        Dim IntializingQuery As String = "Select a,b,Time,[Call Type] from [bts$]"

        Dim QueryOnly_a_party As String = "Select Distinct(a),Time from [bts$] where Time = (Select min(Time) from [bts$])"
        Dim TotalRecordsQuery As String = "Select Count(*) from [bts$]"
        'Dim QueryOnly_a_party As String = "Select * from bts$"
        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & lbGroupSheetPath.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        Conn_bParty = New OleDb.OleDbConnection(ConnectionString)
        Conn_b_to_a = New OleDb.OleDbConnection(ConnectionString)
        Conn_b_to_a.Open()
        cmd_b_to_a = New OleDb.OleDbCommand(IntializingQuery, Conn_b_to_a)
        Reader_b_to_a = cmd_b_to_a.ExecuteReader

        While Reader_b_to_a.Read
            If IsDBNull(Reader_b_to_a(1)) = False Then
                bParty = Reader_b_to_a(1)
            Else
                bParty = ""
            End If
            'Call InsertInAnagramTable(Reader_b_to_a(0), bParty, Reader_b_to_a(2), Reader_b_to_a(3), "tblAnagram")
        End While
        Reader_b_to_a.Close()
        cmd_b_to_a.Dispose()
        Conn_b_to_a.Close()
        ConnCounter_bParty = New OleDb.OleDbConnection(ConnectionString)

        ConnCount_b_to_a = New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        'cmdOnly_a_party = New OleDb.OleDbCommand(IntializingQuery, o)
        cmdOnly_a_party = New OleDb.OleDbCommand(TotalRecordsQuery, o)
        TotalRecords = CInt(cmdOnly_a_party.ExecuteScalar())
        lbTotalNoOfaParty.Text = TotalRecords
        lbTotalNoOfaParty.Refresh()
        ' Dim GroupNumber As Integer = 1
        'lbCheckedaPraty.Text = totalcalls
        If TotalRecords >= 1 Then
            Call CreateVisioDoc()
        End If
        cmdOnly_a_party.Dispose()
        While TotalRecords >= 1
            If is_VFileSave = True Then
                vDoc = vApp.Documents.Add("")
                vDoc.PaperSize = Visio.VisPaperSizes.visPaperSizeLegal
                'vApp.Visible = False
                vStencil = vApp.Documents.OpenEx("Basic Flowchart Shapes.vss", 4)
                vConnectorStencil = vApp.Documents.OpenEx("Connectors.vss", 4)
                vFlowChartMaster = vStencil.Masters("Process")
                vConnectorMaster = vConnectorStencil.Masters("Line connector")
                is_VFileSave = False
            End If
            If pageCounter > 1 Then
                vDoc.Pages.Add()
            End If
            Call TitleOfGroup(pageCounter)
            Y_LS_Shape = 12.0
            Y_RS_Shape = 12.0
            AnagramGroupTable.Clear()
            cmdOnly_a_party = New OleDb.OleDbCommand(QueryOnly_a_party, o)
            Only_a_prtyReader = cmdOnly_a_party.ExecuteReader
            While Only_a_prtyReader.Read
                aParty = Only_a_prtyReader(0).ToString
                InitialDatetime = Only_a_prtyReader(1)
                CallDateTime = Only_a_prtyReader(1)
            End While
            Only_a_prtyReader.Close()
            Only_a_prtyReader = cmdOnly_a_party.ExecuteReader()
            'o.Close()
            cmdOnly_a_party.Dispose()
            extractingQuery = "Select (Select Count(a) as TotalCalls from [bts$] where a = " & aParty & ") ,(Select Count([Call Type]) as Incomings from [bts$] where a = " & aParty & " AND [Call Type] = 'In'),(Select Count([Call Type]) as Incomings from [bts$] where a = " & aParty & " AND [Call Type] = 'Out') from [bts$] where a = " & aParty & ""
            cmdOnly_a_party = New OleDb.OleDbCommand(extractingQuery, o)
            totalcalls_a = 0
            Incommings_a = 0
            Outgoings_a = 0
            Only_a_prtyReader = cmdOnly_a_party.ExecuteReader
            Dim totalIndex As Integer = Nothing
            While Only_a_prtyReader.Read
                totalcalls_a = Only_a_prtyReader(0)
                Incommings_a = Only_a_prtyReader(1)
                Outgoings_a = Only_a_prtyReader(2)
                ShapeText = aParty.Substring(aParty.Length - 10, 10) & "(" & totalcalls_a & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & Incommings_a & ") Out(" & Outgoings_a & ")"
                Exit While
            End While
            lbtotalCallsa.Text = CInt(lbtotalCallsa.Text) + totalcalls_a
            lbtotalCallsa.Refresh()
            Call DropCentralShape()
            'Call InsertToAnagramTable("a", aParty, totalcalls_a, Incommings_a, Outgoings_a, CallDateTime, 1)
            Only_a_prtyReader.Close()
            cmdOnly_a_party.Dispose()
            o.Close()
            Query_bParty = "Select Distinct(b) from [bts$] where a = " & aParty & ""
            Conn_bParty.Open()
            cmd_bParty = New OleDb.OleDbCommand(Query_bParty, Conn_bParty)
            Reader_bParty = cmd_bParty.ExecuteReader
            ShapCounter = 1
            While Reader_bParty.Read
                bParty = Reader_bParty(0).ToString
                ' If IsDBNull(Reader_bParty(0)) = False And bParty.Length >= 10 Then
                If IsDBNull(Reader_bParty(0)) = True Then
                    bParty = "Null Value"
                    QueryCounter_bParty = "Select (Select Count(*) as TotalCalls from [bts$] where a = " & aParty & " AND  b IS NULL) ,(Select Count([Call Type]) as Incomings from [bts$] where a = " & aParty & " AND b IS NULL AND [Call Type] = 'In'),(Select Count([Call Type]) as Incomings from [bts$] where a = " & aParty & " AND b IS NULL AND [Call Type] = 'Out'), (Select max(Time) as EndTime from [bts$] where a = " & aParty & " AND  b IS NULL) from [bts$] where a = " & aParty & " AND b IS NULL"
                Else
                    QueryCounter_bParty = "Select (Select Count(b) as TotalCalls from [bts$] where b = " & bParty & " AND a = " & aParty & ") ,(Select Count([Call Type]) as Incomings from [bts$] where a = " & aParty & " AND b = " & bParty & " AND [Call Type] = 'In'),(Select Count([Call Type]) as Incomings from [bts$] where a = " & aParty & " AND b = " & bParty & " AND [Call Type] = 'Out'), (Select max(Time) as EndTime from [bts$] where b = " & bParty & " AND a = " & aParty & ") from [bts$] where a = " & aParty & " AND b = " & bParty & ""
                End If
                bCounter += 1
                'EndDatetime = Reader_bParty(1)
                'CallDateTime = Reader_bParty(1)

                ConnCounter_bParty.Open()
                cmd_Counter_bPary = New OleDb.OleDbCommand(QueryCounter_bParty, ConnCounter_bParty)
                ReaderCounter_bParty = cmd_Counter_bPary.ExecuteReader
                While ReaderCounter_bParty.Read
                    'Call InsertToAnagramTable("b", bParty, ReaderCounter_bParty(0), ReaderCounter_bParty(1), ReaderCounter_bParty(2), CallDateTime, bCounter)
                    CallTimeResult = DateTime.Compare(CallDateTime, ReaderCounter_bParty(3))
                    If CallTimeResult < 0 Then
                        EndDatetime = ReaderCounter_bParty(3)
                        CallDateTime = ReaderCounter_bParty(3)
                    End If
                    If IsDBNull(Reader_bParty(0)) = True Or bParty.Length < 10 Then
                        ShapeText = bParty & "(" & ReaderCounter_bParty(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & ReaderCounter_bParty(1) & ") Out(" & ReaderCounter_bParty(2) & ")"
                    Else
                        ShapeText = bParty.Substring(bParty.Length - 10, 10) & "(" & ReaderCounter_bParty(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & ReaderCounter_bParty(1) & ") Out(" & ReaderCounter_bParty(2) & ")"
                    End If

                    If CLng(ShapCounter) Mod 2 > 0 Then
                        If IsDBNull(Reader_bParty(0)) = True Or bParty.Length < 10 Then
                            Call dropLeftAtoBShapes(True)
                        Else
                            Call dropLeftAtoBShapes(False)
                        End If

                    Else
                        If IsDBNull(Reader_bParty(0)) = True Or bParty.Length < 10 Then
                            Call dropRightAtoBShapes(True)
                        Else
                            Call dropRightAtoBShapes(False)
                        End If

                    End If
                    Exit While
                End While
                If IsDBNull(Reader_bParty(0)) = False And bParty.Length >= 10 Then
                    If bParty.Length < 10 Then
                        QueryString_b_to_a = "Select Distinct(a)  from [bts$] where b Like '%" & bParty & "'"
                    Else
                        QueryString_b_to_a = "Select Distinct(a)  from [bts$] where b Like '%" & bParty.Substring(bParty.Length - 10, 10) & "'"
                    End If

                    Conn_b_to_a.Open()
                    cmd_b_to_a = New OleDb.OleDbCommand(QueryString_b_to_a, Conn_b_to_a)
                    Reader_b_to_a = cmd_b_to_a.ExecuteReader
                    While Reader_b_to_a.Read
                        bToaParty = Reader_b_to_a(0)
                        ' CallDateTime = Reader_b_to_a(1)
                        QueryStringCount_bToa = "Select (Select Count(a) as TotalCalls from [bts$] where a = " & bToaParty & " AND b Like '%" & bParty.Substring(bParty.Length - 10, 10) & "') ,(Select Count([Call Type]) as Incomings from [bts$] where a = " & bToaParty & " AND [Call Type] = 'In' AND b Like '%" & bParty.Substring(bParty.Length - 10, 10) & "'),(Select Count([Call Type]) as Incomings from [bts$] where a = " & bToaParty & " AND [Call Type] = 'Out' AND b Like '%" & bParty.Substring(bParty.Length - 10, 10) & "'), (Select max(Time) as EndTime from [bts$] where a = " & bToaParty & " AND b Like '%" & bParty.Substring(bParty.Length - 10, 10) & "') from [bts$] where a = " & bToaParty & ""
                        ConnCount_b_to_a.Open()
                        cmdCount_b_to_a = New OleDb.OleDbCommand(QueryStringCount_bToa, ConnCount_b_to_a)
                        ReaderCount_b_to_a = cmdCount_b_to_a.ExecuteReader
                        While ReaderCount_b_to_a.Read
                            CallTimeResult = DateTime.Compare(CallDateTime, ReaderCount_b_to_a(3))
                            If CallTimeResult < 0 Then
                                EndDatetime = ReaderCount_b_to_a(3)
                                CallDateTime = ReaderCount_b_to_a(3)
                            End If
                            'Call InsertToAnagramTable("bToa", bToaParty, ReaderCount_b_to_a(0), ReaderCount_b_to_a(1), ReaderCount_b_to_a(2), CallDateTime, 1)
                            ShapeText = bToaParty.Substring(bToaParty.Length - 10, 10) & "(" & ReaderCount_b_to_a(0) & ")" & vbCrLf & "CNIC" & vbCrLf & "In(" & ReaderCount_b_to_a(1) & ") Out(" & ReaderCount_b_to_a(2) & ")"
                            If CLng(ShapCounter) Mod 2 > 0 Then
                                Call dropLeftBtoAShapes()
                            Else
                                Call dropRightBtoAShapes()
                            End If
                            Exit While
                        End While
                        ReaderCount_b_to_a.Close()
                        cmdCount_b_to_a.Dispose()
                        ConnCount_b_to_a.Close()
                    End While
                    Reader_b_to_a.Close()
                    cmd_b_to_a.Dispose()
                    Conn_b_to_a.Close()
                End If
                ReaderCounter_bParty.Close()
                cmd_Counter_bPary.Dispose()
                ConnCounter_bParty.Close()
                ShapCounter += 1
                'End If

            End While
            Reader_bParty.Close()
            cmd_bParty.Dispose()
            Conn_bParty.Close()
            Call DeleteRecords(aParty)
            o.Open()
            cmdOnly_a_party = New OleDb.OleDbCommand(TotalRecordsQuery, o)
            TotalRecords = CInt(cmdOnly_a_party.ExecuteScalar())
            cmdOnly_a_party.Dispose()
            lbCheckedaPraty.Text = CInt(lbCheckedaPraty.Text) + totalcalls_a
            lbCheckedaPraty.Refresh()
            lbNoOfAnagram.Text = GroupNumber
            lbNoOfAnagram.Refresh()
            If Y_RS_Shape < Y_LS_Shape Then
                aPartyShape.Cells("piny").ResultIU = ((12.0 + Y_RS_Shape) / 2)
            ElseIf Y_LS_Shape < Y_RS_Shape Then
                aPartyShape.Cells("piny").ResultIU = ((12.0 + Y_LS_Shape) / 2)
            Else
                aPartyShape.Cells("piny").ResultIU = ((12.0 + Y_RS_Shape) / 2)
            End If
            vTitleShape.Text = "From:  " & InitialDatetime & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & EndDatetime
            If pageCounter = 25 Then
                vFilePath = vDirectoryPath & "\Anagram of groups on BTS " & pageCounter * vFileCounter + 1 & " to " & GroupNumber & ".vsd"
                Try
                    ' Delete the previous version of the file.
                    Kill(vFilePath)
                Catch
                End Try
                vFileCounter += 1
                vDoc.SaveAs(vFilePath)
                is_VFileSave = True
                pageCounter = 0
                vDoc.Close()

            End If
            pageCounter += 1
            GroupNumber += 1

        End While
        o.Close()
        If is_VFileSave = False Then
            vFilePath = vDirectoryPath & "\Anagram of groups on BTS " & GroupNumber - pageCounter + 1 & " to " & GroupNumber - 1 & ".vsd"
            Try
                ' Delete the previous version of the file.
                Kill(vFilePath)
            Catch
            End Try
            vDoc.SaveAs(vFilePath)
            vDoc.Close()
            vApp.Quit()
            vDoc = Nothing
            vApp = Nothing
        End If
        GroupNumber = Nothing
        ShapeText = Nothing
        vTitleShape = Nothing
        aPartyShape = Nothing
        a_To_bPartyShape = Nothing
        b_To_aPartyShape = Nothing
        vConnector = Nothing
        vpage = Nothing
        vFlowChartMaster = Nothing
        vConnectorMaster = Nothing
        vStencil = Nothing
        vConnectorStencil = Nothing
        vBeginCell = Nothing
        vEndCell = Nothing
        X_Central_Shape = Nothing
        Y_Central_Shape = Nothing
        X_LS_Shape_ab = Nothing
        X_LS_Shape_ba = Nothing
        Y_LS_Shape = Nothing
        X_RS_Shape_ab = Nothing
        X_RS_Shape_ba = Nothing
        Y_RS_Shape = Nothing

        GC.Collect()

    End Sub
    Sub BTS_to_SQL()
        Dim Conn_b_to_a As OleDbConnection
        Dim cmd_b_to_a As OleDbCommand
        Dim Reader_b_to_a As OleDbDataReader
        Dim ConnectionString As String
        Dim CallbParty As String
        Dim CallaParty As String
        Dim CallTime As String
        Dim CallDate As String
        Dim CallType As String
        Dim CallDuration As String
        Dim CallCellID As String
        Dim callIMEI As String
        Dim IMSI As String
        Dim Site As String
        Dim CNIC_a As String
        Dim CNIC_b As String
        Dim isWithCNIC As Boolean = False
        Dim IntializingQuery As String
        'If lbGroupSheetPath.Text.Substring(lbGroupSheetPath.Text.IndexOf(".") - 4, 4) = "CNIC" Then
        IntializingQuery = "Select a,CNIC_a,b,CNIC_b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [bts$]"
        '    isWithCNIC = True
        'Else
        '    IntializingQuery = "Select a,b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [bts$]"
        '    isWithCNIC = False
        'End If
        ' Dim IntializingQuery As String = "Select a,CNIC_a,b,CNIC_b,Time,Date,[Call Type],Duration,[Cell ID],IMEI,IMSI,Site from [bts$]"
        ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & lbGroupSheetPath.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Conn_b_to_a = New OleDb.OleDbConnection(ConnectionString)
        Conn_b_to_a.Open()
        cmd_b_to_a = New OleDb.OleDbCommand(IntializingQuery, Conn_b_to_a)
        Reader_b_to_a = cmd_b_to_a.ExecuteReader
        While Reader_b_to_a.Read
            'If isWithCNIC = True Then
            If IsDBNull(Reader_b_to_a(1)) = False Then
                CNIC_a = Reader_b_to_a(1)
            Else
                CNIC_a = ""
            End If
            If IsDBNull(Reader_b_to_a(2)) = False Then
                CallbParty = Reader_b_to_a(2)
            Else
                CallbParty = ""
            End If
            If IsDBNull(Reader_b_to_a(3)) = False Then
                CNIC_b = Reader_b_to_a(3)
            Else
                CNIC_b = ""
            End If
            If IsDBNull(Reader_b_to_a(7)) = False Then
                CallDuration = Reader_b_to_a(7)
            Else
                CallDuration = ""
            End If
            If IsDBNull(Reader_b_to_a(8)) = False Then
                CallCellID = Reader_b_to_a(8)
            Else
                CallCellID = ""
            End If
            If IsDBNull(Reader_b_to_a(9)) = False Then
                callIMEI = Reader_b_to_a(9)
            Else
                callIMEI = ""
            End If
            If IsDBNull(Reader_b_to_a(10)) = False Then
                IMSI = Reader_b_to_a(10)
            Else
                IMSI = ""
            End If
            If IsDBNull(Reader_b_to_a(11)) = False Then
                Site = Reader_b_to_a(11)
            Else
                Site = ""
            End If

            Call InsertInAnagramTable(Reader_b_to_a(0), CNIC_a, CallbParty, CNIC_b, Reader_b_to_a(4), Reader_b_to_a(5), Reader_b_to_a(6), CallDuration, CallCellID, callIMEI, IMSI, Site, "tblAnagram")
            'Else

            '    CNIC_a = ""

            '    If IsDBNull(Reader_b_to_a(1)) = False Then
            '        CallbParty = Reader_b_to_a(1)
            '    Else
            '        CallbParty = ""
            '    End If

            '    CNIC_b = ""

            '    If IsDBNull(Reader_b_to_a(5)) = False Then
            '        CallDuration = Reader_b_to_a(5)
            '    Else
            '        CallDuration = ""
            '    End If
            '    If IsDBNull(Reader_b_to_a(6)) = False Then
            '        CallCellID = Reader_b_to_a(6)
            '    Else
            '        CallCellID = ""
            '    End If
            '    If IsDBNull(Reader_b_to_a(7)) = False Then
            '        callIMEI = Reader_b_to_a(7)
            '    Else
            '        callIMEI = ""
            '    End If
            '    If IsDBNull(Reader_b_to_a(8)) = False Then
            '        IMSI = Reader_b_to_a(8)
            '    Else
            '        IMSI = ""
            '    End If
            '    If IsDBNull(Reader_b_to_a(9)) = False Then
            '        Site = Reader_b_to_a(9)
            '    Else
            '        Site = ""
            '    End If
            '    Call InsertInAnagramTable(Reader_b_to_a(0), CNIC_a, CallbParty, CNIC_b, Reader_b_to_a(2), Reader_b_to_a(3), Reader_b_to_a(4), CallDuration, CallCellID, callIMEI, IMSI, Site, "tblAnagram")
            'End If

        End While
    End Sub
    Sub CreateAnagramTable(ByVal tblAnagram As String)

        Try
            Call ConnectionOpen()
            QueryString = "IF OBJECT_ID('dbo." & tblAnagram & "') IS NOT NULL DROP TABLE " & tblAnagram & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
        Catch ex As Exception
            If TargetConnection.State <> ConnectionState.Closed Then
                TargetConnection.Close()
            End If
        End Try
        Call ConnectionOpen()
        Try
            QueryString = "CREATE TABLE " & tblAnagram & "(a varchar(100),CNIC_a text, b varchar(100), CNIC_b text, Time varchar(100), Date varchar(100), [Call Type] varchar(50), Duration varchar(30),[Cell ID] varchar(100), IMEI varchar(100), IMSI varchar(100), Site text)"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
            TargetConnection.Dispose()
            CreateTbCommand.Dispose()
        Catch ex As Exception

        End Try

    End Sub
    Sub InsertInAnagramTable(ByVal a As String, ByVal CNIC_a As String, ByVal b As String, ByVal CNIC_b As String, ByVal Time As DateTime, ByVal GrpDate As DateTime, ByVal CallType As String, ByVal Duration As String, ByVal CellID As String, ByVal IMEI As String, ByVal IMSI As String, ByVal Site As String, ByVal TableName As String)
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO [" & TableName & "] VALUES ('" & a & "','" & CNIC_a & "','" & b & "','" & CNIC_b & "','" & Time & "','" & GrpDate & "','" & CallType & "','" & Duration & "','" & CellID & "','" & IMEI & "','" & IMSI & "','" & Site & "')"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        Try
            InsertIMEICommand.ExecuteNonQuery()
        Catch ex1 As Exception
            MsgBox("Error in InsertInAnagramTable" & Err.Description & vbCrLf & "Error During Executing Query" & InsertQuery, MsgBoxStyle.Information)
        End Try

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub createGrpLayer(ByVal layerName As String)
        Call ConnectionOpen()
        Try
            QueryString = "CREATE TABLE " & layerName & "(Parent varchar(16), Child varchar(16), TotalCalls varchar(6), Incomings varchar(6), Outgoings varchar(6))"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
            TargetConnection.Dispose()
            CreateTbCommand.Dispose()
        Catch ex As Exception

        End Try
    End Sub
    Sub InsertGrpLayer(ByVal layerName As String, ByVal Parent As String, ByVal Child As String, ByVal TotalCalls As String, ByVal Incomings As String, ByVal Outgoings As String)
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO [" & layerName & "] VALUES ('" & Parent & "','" & Child & "','" & TotalCalls & "','" & Incomings & "','" & Outgoings & "')"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        InsertIMEICommand.ExecuteNonQuery()
        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub DropGrpLayer(ByRef layerName As String)
        Try
            Call ConnectionOpen()
            QueryString = "IF OBJECT_ID('dbo." & layerName & "') IS NOT NULL DROP TABLE " & layerName & ""
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
        Catch ex As Exception
            If TargetConnection.State <> ConnectionState.Closed Then
                TargetConnection.Close()
            End If
        End Try
    End Sub
    Sub DeleteRecords(ByVal PhoneNumber As String)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        xlWorkbook = xlWorkbooks.Open(lbGroupSheetPath.Text)
        For Each xlSheet In xlWorkbook.Worksheets
            If xlSheet.Name = "bts" Then
                xlSheet = xlWorkbook.Worksheets("bts")
                IsSheetRenamed = True
                Exit For
            End If
        Next xlSheet
        Dim lastRow As Long
        Dim j As Long
        With xlSheet
            lastRow = .Range("a" & .Rows.Count).End(Excel.XlDirection.xlUp).Row 'assumes every row in column A has a value
            For j = lastRow To 2 Step -1
                If LCase(.Range("a" & j).Value) = PhoneNumber Then
                    .Range("a" & j).EntireRow.Delete()
                End If
            Next j
        End With
        xlApp.DisplayAlerts = False
        xlWorkbook.SaveAs(lbGroupSheetPath.Text, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkbooks = Nothing
        xlWorkbook.Close(True, misValue, misValue)
        xlApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) : xlWorkbook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet) : xlSheet = Nothing
    End Sub
    Dim AddRow As DataRow

    Sub InsertToAnagramTable(ByVal Party As String, ByVal Number As String, ByVal TotalCalls As String, ByVal IncomingCalls As String, ByVal OutgoingCalls As String, ByVal DateAndTime As Date, ByVal Counter As Integer)
        AddRow = AnagramGroupTable.NewRow
        AddRow(0) = Party
        AddRow(1) = Number
        AddRow(2) = TotalCalls
        AddRow(3) = IncomingCalls
        AddRow(4) = OutgoingCalls
        AddRow(5) = DateAndTime
        AddRow(6) = Counter
        AnagramGroupTable.Rows.Add(AddRow)
    End Sub
    Dim AnagramGroupTable As DataTable
    Dim ds As DataSet

    Sub CreateAnagramTable()
        AnagramGroupTable = New DataTable
        AnagramGroupTable.Columns.Add("Party", Type.GetType("System.String"))
        AnagramGroupTable.Columns.Add("Number", Type.GetType("System.String"))
        AnagramGroupTable.Columns.Add("TotalCalls", Type.GetType("System.String"))
        AnagramGroupTable.Columns.Add("IncomingCalls", Type.GetType("System.String"))
        AnagramGroupTable.Columns.Add("OutgoingCalls", Type.GetType("System.String"))
        AnagramGroupTable.Columns.Add("DateAndTiem", Type.GetType("System.DateTime"))
        AnagramGroupTable.Columns.Add("Counter", Type.GetType("System.Int32"))
    End Sub
    Dim vApp As Visio.Application
    Dim vDoc As Visio.Document
    Dim isVdocCreated As Boolean = False

    Sub CreateVisioDoc()
        Try


            vApp = New Visio.Application
            vApp.Visible = True
            'Create a new document; note the empty string.
            vDoc = vApp.Documents.Add("")
            vDoc.PaperSize = Visio.VisPaperSizes.visPaperSizeLegal
            vApp.Visible = True
            vStencil = vApp.Documents.OpenEx("Basic Flowchart Shapes.vss", 4)
            vConnectorStencil = vApp.Documents.OpenEx("Connectors.vss", 4)
            vFlowChartMaster = vStencil.Masters("Process")
            vConnectorMaster = vConnectorStencil.Masters("Side to Side 1")
            isVdocCreated = True
        Catch ex1 As Exception
            MsgBox(Err.Description & "During creating visio Doc")
        End Try
    End Sub
    Dim GroupNumber As Integer = 1
    Dim ShapeText As String = Nothing
    Dim vTitleShape As Visio.Shape
    Dim aPartyShape As Visio.Shape
    Dim LeftShape() As Visio.Shape
    Dim RightShape() As Visio.Shape
    Dim a_To_bPartyShape As Visio.Shape
    Dim b_To_aPartyShape As Visio.Shape
    Dim vConnector As Visio.Shape
    Dim vpage As Visio.Page
    Dim vFlowChartMaster As Visio.Master
    Dim vConnectorMaster As Visio.Master
    Dim vStencil As Visio.Document
    Dim vConnectorStencil As Visio.Document
    Dim vBeginCell As Visio.Cell
    Dim vEndCell As Visio.Cell
    Dim X_Central_Shape As Double = 4.25
    Dim Y_Central_Shape As Double = 12.0
    Dim X_LS_Shape_ab As Double = 2.15
    Dim X_LS_Shape_ba As Double = 0.75
    Dim Y_LS_Shape As Double = 12.0
    Dim X_RS_Shape_ab As Double = 6.45
    Dim X_RS_Shape_ba As Double = 7.75
    Dim Y_RS_Shape As Double = 12.0
    Sub GenerateAnagram()
        'Dim ShapeText As String = Nothing
        'Dim vTitleShape As Visio.Shape
        'Dim aPartyShape As Visio.Shape
        'Dim a_To_bPartyShape As Visio.Shape
        'Dim b_To_aPartyShape As Visio.Shape
        'Dim vConnector As Visio.Shape
        'Dim vpage As Visio.Page
        'Dim vFlowChartMaster As Visio.Master
        'Dim vConnectorMaster As Visio.Master
        'Dim vStencil As Visio.Document
        'Dim vConnectorStencil As Visio.Document
        'Dim vBeginCell As Visio.Cell
        'Dim vEndCell As Visio.Cell
        'Dim X_Central_Shape As Double = 4.25
        'Dim Y_Central_Shape As Double = 8
        'Dim X_LS_Shape_ab As Double
        'Dim X_LS_Shape_ba As Double
        'Dim Y_LS_Shape As Double
        'Dim X_RS_Shape_ab As Double
        'Dim X_RS_Shape_ba As Double
        'Dim Y_RS_Shape As Double
        'Dim Result() As DataRow
        'If GroupNumber > 1 Then
        '    vDoc.Pages.Add()
        'End If
        'vStencil = vApp.Documents.OpenEx("Basic Flowchart Shapes.vss", 4)
        'vConnectorStencil = vApp.Documents.OpenEx("Connectors.vss", 4)
        'vpage = vDoc.Pages(GroupNumber)
        'vpage.PageSheet.Cells("PageHeight").ResultIU = 14
        'vpage.PageSheet.Cells("PageWidth").ResultIU = 8.5
        'vDoc.PaperSize = Visio.VisPaperSizes.visPaperSizeLegal
        'vpage.Name = "Group" & GroupNumber

        ''GroupBox Title
        'vFlowChartMaster = vStencil.Masters("Process")
        'vTitleShape = vApp.ActivePage.Drop(vFlowChartMaster, 4.25, 13.55)
        'vTitleShape.Cells("Char.Size").Result("pt") = 20
        'vTitleShape.Characters.CharProps(2) = &H1
        'vTitleShape.CellsU("Fillforegnd").Formula = "RGB(255, 235, 205)"
        'vTitleShape.Cells("Width").ResultIU = 8
        'vTitleShape.Cells("Height").ResultIU = 0.35
        'vTitleShape.Text = "Anagram of groups on BTS"

        ''Group Number
        'vFlowChartMaster = vStencil.Masters("Process")
        'vTitleShape = vApp.ActivePage.Drop(vFlowChartMaster, 4.25, 13.17)
        'vTitleShape.Cells("Char.Size").Result("pt") = 14
        'vTitleShape.Characters.CharProps(2) = &H1
        'vTitleShape.CellsU("Fillforegnd").Formula = "RGB(255, 235, 205)"
        'vTitleShape.Cells("Width").ResultIU = 4
        'vTitleShape.Cells("Height").ResultIU = 0.25
        'vTitleShape.Text = "Group" & GroupNumber

        ''Start Date and End Date
        'vFlowChartMaster = vStencil.Masters("Process") '(aryValues(iCount, 0))
        'vTitleShape = vApp.ActivePage.Drop(vFlowChartMaster, 4.25, 12.82)
        'vTitleShape.Cells("Char.Size").Result("pt") = 12
        'vTitleShape.Characters.CharProps(2) = &H1
        'vTitleShape.CellsU("Fillforegnd").Formula = "RGB(255, 235, 205)"
        'vTitleShape.Cells("Width").ResultIU = 8
        'vTitleShape.Cells("Height").ResultIU = 0.25
        'vTitleShape.Text = "From:  " & Date.Now & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & Date.Now
        'Result = AnagramGroupTable.Select("Party='a'")
        'For Each row As DataRow In Result
        '    '"In(9) Out(0)"
        '    ShapeText = row(1) & "(" & row(2) & ")" & vbCrLf & "IN(" & row(3) & ") Out(" & row(4) & ")"
        'Next
        'vFlowChartMaster = vStencil.Masters("Process")
        'aPartyShape = vApp.ActivePage.Drop(vFlowChartMaster, X_Central_Shape, Y_Central_Shape)
        'aPartyShape.Characters.CharProps(2) = &H1
        'aPartyShape.CellsU("Fillforegnd").Formula = "RGB(205, 205, 193)"
        'aPartyShape.Cells("Width").ResultIU = 0.95
        'aPartyShape.Cells("Height").ResultIU = 0.45
        'aPartyShape.Text = ShapeText
        'Result = AnagramGroupTable.Select("Party='ab'")
        'For Each row As DataRow In Result

        'Next
    End Sub
    Sub TitleOfGroup(ByVal PageNumber As Integer)
        vpage = vDoc.Pages(PageNumber)
        vpage.PageSheet.Cells("PageHeight").ResultIU = 14
        vpage.PageSheet.Cells("PageWidth").ResultIU = 8.5
        vDoc.PaperSize = Visio.VisPaperSizes.visPaperSizeLegal
        vpage.Name = "Group" & GroupCounter

        'Title of group

        vTitleShape = vApp.ActivePage.Drop(vFlowChartMaster, 4.25, 13.55)
        vTitleShape.Cells("Char.Size").Result("pt") = 20
        vTitleShape.Characters.CharProps(2) = &H1
        vTitleShape.CellsU("Fillforegnd").Formula = "RGB(255, 235, 205)"
        vTitleShape.Cells("Width").ResultIU = 8
        vTitleShape.Cells("Height").ResultIU = 0.35
        vTitleShape.Text = "Anagram of groups on BTS"

        ' Number of Group
        'vFlowChartMaster = vStencil.Masters("Process")
        vTitleShape = vApp.ActivePage.Drop(vFlowChartMaster, 4.25, 13.17)
        vTitleShape.Cells("Char.Size").Result("pt") = 14
        vTitleShape.Characters.CharProps(2) = &H1
        vTitleShape.CellsU("Fillforegnd").Formula = "RGB(255, 235, 205)"
        vTitleShape.Cells("Width").ResultIU = 4
        vTitleShape.Cells("Height").ResultIU = 0.25
        vTitleShape.Text = "Group" & GroupCounter

        'Start Date and End Date
        'vFlowChartMaster = vStencil.Masters("Process")
        vTitleShape = vApp.ActivePage.Drop(vFlowChartMaster, 4.25, 12.82)
        vTitleShape.Cells("Char.Size").Result("pt") = 12
        vTitleShape.Characters.CharProps(2) = &H1
        vTitleShape.CellsU("Fillforegnd").Formula = "RGB(255, 235, 205)"
        vTitleShape.Cells("Width").ResultIU = 8
        vTitleShape.Cells("Height").ResultIU = 0.25
        vTitleShape.Text = "From:  " & Date.Now & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "To: " & Date.Now
    End Sub
    Sub DropCentralShape()
        'vFlowChartMaster = vStencil.Masters("Process")
        aPartyShape = vApp.ActivePage.Drop(vFlowChartMaster, X_Central_Shape, Y_Central_Shape)
        aPartyShape.Characters.CharProps(2) = &H1
        aPartyShape.CellsU("Fillforegnd").Formula = "RGB(238, 221, 130)"
        aPartyShape.Cells("Width").ResultIU = 0.95
        aPartyShape.Cells("Height").ResultIU = 0.45
        aPartyShape.Text = ShapeText
        aPartyShape.Name = ShapeText.Substring(0, ShapeText.IndexOf("(")) & "Central"
        'aPartyShape.Cells("piny").ResultIU = Y_Central_Shape + 1.5
    End Sub

    Sub dropLeftAtoBShapes(ByVal EmptyValue As Boolean)
        a_To_bPartyShape = vApp.ActivePage.Drop(vFlowChartMaster, X_LS_Shape_ab, Y_LS_Shape)
        ReDim Preserve LeftShape(1)

        a_To_bPartyShape.Characters.CharProps(2) = &H1
        a_To_bPartyShape.CellsU("Fillforegnd").Formula = "RGB(255, 228, 225)"
        a_To_bPartyShape.Cells("Width").ResultIU = 0.95
        a_To_bPartyShape.Cells("Height").ResultIU = 0.45
        a_To_bPartyShape.Text = ShapeText
        Dim mytext As String = a_To_bPartyShape.Text
        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(aPartyShape.Cells("AlignLeft"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(a_To_bPartyShape.Cells("AlignRight"))
        If EmptyValue = True Then
            Y_LS_Shape -= 0.55
        End If
    End Sub
    Sub dropLeftBtoAShapes()
        b_To_aPartyShape = vApp.ActivePage.Drop(vFlowChartMaster, X_LS_Shape_ba, Y_LS_Shape)
        b_To_aPartyShape.Characters.CharProps(2) = &H1
        b_To_aPartyShape.CellsU("Fillforegnd").Formula = "RGB(224, 255, 255)"
        b_To_aPartyShape.Cells("Width").ResultIU = 0.95
        b_To_aPartyShape.Cells("Height").ResultIU = 0.45
        b_To_aPartyShape.Text = ShapeText
        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(a_To_bPartyShape.Cells("AlignLeft"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(b_To_aPartyShape.Cells("AlignRight"))
        Y_LS_Shape -= 0.55
    End Sub
    Sub dropRightAtoBShapes(ByVal EmptyValue As Boolean)
        a_To_bPartyShape = vApp.ActivePage.Drop(vFlowChartMaster, X_RS_Shape_ab, Y_RS_Shape)
        a_To_bPartyShape.Characters.CharProps(2) = &H1
        a_To_bPartyShape.CellsU("Fillforegnd").Formula = "RGB(255, 228, 225)"
        a_To_bPartyShape.Cells("Width").ResultIU = 0.95
        a_To_bPartyShape.Cells("Height").ResultIU = 0.45
        a_To_bPartyShape.Text = ShapeText
        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(aPartyShape.Cells("AlignRight"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(a_To_bPartyShape.Cells("AlignLeft"))
        If EmptyValue = True Then
            Y_RS_Shape -= 0.55
        End If
    End Sub
    Sub dropRightBtoAShapes()
        b_To_aPartyShape = vApp.ActivePage.Drop(vFlowChartMaster, X_RS_Shape_ba, Y_RS_Shape)
        b_To_aPartyShape.Characters.CharProps(2) = &H1
        b_To_aPartyShape.CellsU("Fillforegnd").Formula = "RGB(224, 255, 255)"
        b_To_aPartyShape.Cells("Width").ResultIU = 0.95
        b_To_aPartyShape.Cells("Height").ResultIU = 0.45
        b_To_aPartyShape.Text = ShapeText
        vConnector = vApp.ActivePage.Drop(vConnectorMaster, 0, 0)
        vBeginCell = vConnector.Cells("BeginY")
        vBeginCell.GlueTo(a_To_bPartyShape.Cells("AlignRight"))
        vEndCell = vConnector.Cells("EndY")
        vEndCell.GlueTo(b_To_aPartyShape.Cells("AlignLeft"))
        Y_RS_Shape -= 0.55
    End Sub
    Sub SaveVisioFile(ByVal vFilePath As String)

    End Sub

    
End Class
