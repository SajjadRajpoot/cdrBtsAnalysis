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

Module mdlExcelManipulation
    Function CommonNos(ByVal All_Files() As String)
        ''Public AllCommonFileNames() As String
        '' Public OnlyFilesPath As String
        Dim TotalNumberOfCDRs As Integer = All_Files.Length()
        Dim OnlyFileNames(TotalNumberOfCDRs) As String
        Dim TempTableName As String
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim SheetName As String = "cdr$"
        Dim InsertCommand As OleDb.OleDbCommand
        Dim ds As New DataSet
        Dim dtQuery As String
        'Dim dt As System.Data.DataTable
        'For j As Integer = 0 To TotalNumberOfCDRs - 1
        '    OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(All_Files(j))
        '    TempTableName = "_" & OnlyFileNames(j)
        '    Dim column1 As DataColumn = New DataColumn(TempTableName)
        '    column1.DataType = System.Type.GetType("System.String")
        '    dt.Columns.Add(column1)
        'Next

        For j As Integer = 0 To TotalNumberOfCDRs - 1

            OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(All_Files(j))
            TempTableName = "_" & OnlyFileNames(j)

            Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & All_Files(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
            Dim o As New OleDb.OleDbConnection(ConnectionString)
            o.Open()
            'populate table with data with 92
            Dim queryString As String = "select party from (select [" & ColumnName1 & "] as Party from [" & SheetName & "]" + _
                        "UNION ALL select [" & ColumnName2 & "] as Party from [" & SheetName & "])"

            InsertCommand = New OleDb.OleDbCommand(queryString, o)
            Dim Da As New OleDbDataAdapter
            Da.SelectCommand = InsertCommand
            Dim dt As New System.Data.DataTable(TempTableName)
            'Dim Column1 As DataColumn = New DataColumn(TempTableName & "_")
            'Column1.DataType = System.Type.GetType("System.String")
            dt.Clear()
            Da.Fill(dt)
            ds.Tables.Add(dt)
            o.Close()
            dtQuery = "party = '921506797288'"
            Dim count As Integer '= ds.Tables(TempTableName).Rows.Count
            Dim row() As DataRow = ds.Tables(0).Select(dtQuery)

            'Dim row() As DataRow = ds.Tables(0).Rows.Find(921506797288)
            count = row.Count

        Next
    End Function

    Function Excel_To_SQL(ByVal AllCommonFileNames() As String, ByVal onlyXLFilePath As String)
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim SheetName As String = "cdr$"
        Dim TotalNumberOfCDRs As Integer = AllCommonFileNames.Length()
        Dim OnlyFileNames(TotalNumberOfCDRs) As String
        frm_Spy_Tech.prgbar_Common_Links.Minimum = 0
        frm_Spy_Tech.prgbar_Common_Links.Maximum = AllCommonFileNames.Length
        frm_Spy_Tech.prgbar_Common_Links.Value = 0
        frm_Spy_Tech.prgbar_Common_Links.Visible = True
        frm_Spy_Tech.prgbar_Common_Links.Refresh()
        Dim TransferedFiles As Integer = 0
        frm_Spy_Tech.lbCommonNos.Text = "Excel to SQL " & TransferedFiles.ToString & " of " & TotalNumberOfCDRs.ToString
        frm_Spy_Tech.lbCommonNos.Visible = True
        frm_Spy_Tech.lbCommonNos.Refresh()
        Dim db2021Connection As SqlConnection

        'Create DataTable for results
        OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
        OthersConnection.Open()
        Dim tempdbConn As SqlConnection = New SqlConnection("Server=" + ServerName + ";Database=tempdb;Trusted_Connection=True;")
        tempdbConn.Open()
        Dim TargetConnection As SqlConnection
        TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=tempdb;Trusted_Connection=True;")
        TargetConnection.Open()
        db2021Connection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;")
        db2021Connection.Open()
        'Dim Delquery As String = "DELETE FROM CommonNumbers"
        'OthersCommand = New SqlCommand(Delquery, OthersConnection)
        'OthersCommand.ExecuteNonQuery()
        Dim CreateTablesQuery As String
        Dim TempTableName As String
        Dim QueryInsert As String
        Dim queryString1 As String
        Dim CreateTbCommand As SqlCommand

        Dim ConnectionString As String
        Dim o As OleDb.OleDbConnection
        Dim queryString As String
        Dim InsertCommand As OleDb.OleDbCommand
        Dim tblCommonNumbers As String = "[CommonNumbers] nvarchar(255) null,"

        Dim DtColName As New Data.DataTable
        DtColName.Clear()
        DtColName.Columns.Add("ColNames", GetType(System.String))
        For j As Integer = 0 To TotalNumberOfCDRs - 1
            OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))

            TempTableName = "_" & OnlyFileNames(j)
            TempTableName = TempTableName
            DtColName.Rows.Add(TempTableName)
            Try
                queryString1 = "Drop Table IF EXISTS [" & TempTableName & "]"
                CreateTbCommand = New SqlCommand(queryString1, tempdbConn)
                CreateTbCommand.ExecuteNonQuery()
            Catch ex03 As Exception
            End Try
            CreateTablesQuery = "CREATE TABLE [" & TempTableName & "] (PhoneNumber nvarchar(255) null)"
            tblCommonNumbers = tblCommonNumbers & "[" & TempTableName & "] nvarchar(255) null,"
            OthersCommand = New SqlCommand(CreateTablesQuery, tempdbConn)
            Try
                OthersCommand.ExecuteNonQuery()
                'TargetConnection.Close()
            Catch ex02 As Exception
                MsgBox("creating table", MsgBoxStyle.OkOnly)
                'TargetConnection.Close()
            End Try
            'CreateTablesQuery = "CREATE CLUSTERED INDEX myIdx ON  [" & TempTableName & "](PhoneNumber)"
            'OthersCommand = New SqlCommand(CreateTablesQuery, OthersConnection)
            'Try
            '    OthersCommand.ExecuteNonQuery()
            'Catch ex02 As Exception
            '    MsgBox("creating table", MsgBoxStyle.OkOnly)
            'End Try
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
            o = New OleDb.OleDbConnection(ConnectionString)
            o.Open()
            'populate table with data with 92
            queryString = "select Party from (select [" & ColumnName1 & "] as Party from [" & SheetName & "] WHERE Isnumeric(" & ColumnName1 & ") AND LEN(" & ColumnName1 & ")>=10) " + _
                        "UNION ALL (select [" & ColumnName2 & "] as Party from [" & SheetName & "] WHERE Isnumeric(" & ColumnName2 & ") AND LEN(" & ColumnName2 & ")>=10)"
            InsertCommand = New OleDb.OleDbCommand(queryString, o)
            Dim InsertReader As OleDb.OleDbDataReader
            Try
                InsertReader = InsertCommand.ExecuteReader()
            Catch ex02 As Exception
                MsgBox("insertion error table", MsgBoxStyle.OkOnly)
            End Try
            Dim bcCopy As New SqlBulkCopy(TargetConnection)
            'TargetConnection.Open()
            bcCopy.BatchSize = 100000
            bcCopy.BulkCopyTimeout = 0
            bcCopy.DestinationTableName = "[" & TempTableName & "]"
            bcCopy.WriteToServer(InsertReader)
            InsertReader.Close()
            'strConSrc.Close()
            'TargetConnection.Close()
            'End While
            'InsertReader.Close()
            o.Close()
            TransferedFiles = TransferedFiles + 1
            frm_Spy_Tech.prgbar_Common_Links.Value = frm_Spy_Tech.prgbar_Common_Links.Value + 1
            'frm_Spy_Tech.prgbar_Common_Links.Refresh()
            frm_Spy_Tech.lbCommonNos.Text = "Chaching the data " & TransferedFiles & " of " & TotalNumberOfCDRs
            frm_Spy_Tech.Refresh()
        Next
        'OthersConnection.Close()
        'TargetConnection.Close()
        'Now Scanning common numbers
        Try
            queryString1 = "Drop Table IF EXISTS [CommonNumbers]"
            CreateTbCommand = New SqlCommand(queryString1, db2021Connection)
            CreateTbCommand.ExecuteNonQuery()
        Catch ex03 As Exception
        End Try

        tblCommonNumbers = "CREATE TABLE [CommonNumbers] (" & tblCommonNumbers.Substring(0, tblCommonNumbers.Length - 1) & ")"
        OthersCommand = New SqlCommand(tblCommonNumbers, db2021Connection)
        Try
            OthersCommand.ExecuteNonQuery()
            'TargetConnection.Close()
        Catch ex02 As Exception
            MsgBox("Error creating Common Numbers table", MsgBoxStyle.OkOnly)
            'TargetConnection.Close()
        End Try
        'Dim DeleteQuery As String = "Delete"
        Dim updateQuery As String = Nothing
        TotalNumberOfCDRs = DtColName.Rows.Count
        TransferedFiles = 0
        frm_Spy_Tech.prgbar_Common_Links.Maximum = TotalNumberOfCDRs
        frm_Spy_Tech.prgbar_Common_Links.Value = 0
        frm_Spy_Tech.lbCommonNos.Text = "Making data consistance " & TransferedFiles & " of " & TotalNumberOfCDRs
        frm_Spy_Tech.Refresh()
        If DtColName.Rows.Count > 0 Then
            For i As Integer = 0 To DtColName.Rows.Count - 1
                TransferedFiles = i + 1
                updateQuery = "update [" + DtColName.Rows(i)(0).ToString + "] set [" + DtColName.Rows(i)(0).ToString + "].PhoneNumber='92'+Substring([" + DtColName.Rows(i)(0).ToString + "].PhoneNumber,2,LEN([" + DtColName.Rows(i)(0).ToString + "].PhoneNumber)) " + _
                   " where [" + DtColName.Rows(i)(0).ToString + "].PhoneNumber like '0%' " + _
                   "AND [" + DtColName.Rows(i)(0).ToString + "].PhoneNumber NOT like   '00%'"
                OthersCommand = New SqlCommand(updateQuery, tempdbConn)
                OthersCommand.ExecuteNonQuery()
                updateQuery = "update [" + DtColName.Rows(i)(0).ToString + "] set [" + DtColName.Rows(i)(0).ToString + "].PhoneNumber='92'+[" + DtColName.Rows(i)(0).ToString + "].PhoneNumber " + _
                   " where [" + DtColName.Rows(i)(0).ToString + "].PhoneNumber like '3%' "

                OthersCommand = New SqlCommand(updateQuery, tempdbConn)
                OthersCommand.ExecuteNonQuery()
                frm_Spy_Tech.prgbar_Common_Links.Value = frm_Spy_Tech.prgbar_Common_Links.Value + 1
                'frm_Spy_Tech.prgbar_Common_Links.Refresh()
                frm_Spy_Tech.lbCommonNos.Text = "Making data consistance " & TransferedFiles & " of " & TotalNumberOfCDRs
                frm_Spy_Tech.Refresh()
            Next
        End If
        Try
            queryString1 = "Drop Table IF EXISTS [CommonNos]"
            CreateTbCommand = New SqlCommand(queryString1, tempdbConn)
            CreateTbCommand.ExecuteNonQuery()
        Catch ex03 As Exception
        End Try

        tblCommonNumbers = "CREATE TABLE [CommonNos] (PhoneNumber nvarchar(255) null)"
        OthersCommand = New SqlCommand(tblCommonNumbers, tempdbConn)
        Try
            OthersCommand.ExecuteNonQuery()
            'TargetConnection.Close()
        Catch ex02 As Exception
            MsgBox("Error creating Common Numbers table", MsgBoxStyle.OkOnly)
            'TargetConnection.Close()
        End Try

        'Copy unique numbers of all files in CommonNos table in sql server

        frm_Spy_Tech.prgbar_Common_Links.Maximum = TotalNumberOfCDRs
        frm_Spy_Tech.prgbar_Common_Links.Value = 0
        frm_Spy_Tech.lbCommonNos.Text = "Copying unique numbers " & TransferedFiles & " of " & TotalNumberOfCDRs
        frm_Spy_Tech.Refresh()
        Dim ColName As String = ""
        Dim SelectQueryRefine As String = ""
        Dim FromQueryRefine As String = ""
        Dim WhereQueryRefine As String = ""
        If DtColName.Rows.Count > 0 Then
            Dim TotalCols As Integer = DtColName.Rows.Count
            For i As Integer = 0 To TotalCols - 1
                ColName = DtColName.Rows(i)(0).ToString
                TransferedFiles = i + 1
                updateQuery = "Select Distinct PhoneNumber from [" + DtColName.Rows(i)(0).ToString + "]"
                OthersCommand = New SqlCommand(updateQuery, tempdbConn)
                Dim InsertReader As SqlDataReader
                InsertReader = OthersCommand.ExecuteReader
                Dim bcCopy As New SqlBulkCopy(TargetConnection)
                bcCopy.BatchSize = 100000
                bcCopy.BulkCopyTimeout = 0
                bcCopy.DestinationTableName = "[CommonNos]"
                bcCopy.WriteToServer(InsertReader)
                InsertReader.Close()
                SelectQueryRefine = SelectQueryRefine + ",(Select Count(" + ColName + ".PhoneNumber) From " + ColName + " Where CommonNos.PhoneNumber=" + ColName + ".PhoneNumber) As " + ColName
                FromQueryRefine = FromQueryRefine + "Left Join " + ColName + " ON " + ColName + ".PhoneNumber=CommonNos.PhoneNumber "
                For j As Integer = i To TotalCols - 1

                    If i <> TotalCols - 1 Then
                        If j <> i Then
                            If i = TotalCols - 2 Then
                                WhereQueryRefine = WhereQueryRefine + "(" + ColName + ">0 and " + DtColName.Rows(j)(0).ToString() + ">0)"
                            Else
                                WhereQueryRefine = WhereQueryRefine + "(" + ColName + ">0 and " + DtColName.Rows(j)(0).ToString() + ">0) OR "
                            End If

                        End If
                        'Else
                        '    If j = i Then
                        '        WhereQueryRefine = WhereQueryRefine + "(" + ColName + ">0 and " + DtColName.Rows(j)(0).ToString() + ">0)"
                        '    End If
                    End If

                Next
                frm_Spy_Tech.prgbar_Common_Links.Value = frm_Spy_Tech.prgbar_Common_Links.Value + 1
                'frm_Spy_Tech.prgbar_Common_Links.Refresh()
                frm_Spy_Tech.lbCommonNos.Text = "Copying unique numbers " & TransferedFiles & " of " & TotalNumberOfCDRs
                frm_Spy_Tech.Refresh()
            Next
        End If

        Dim RefineCommonNosQuery As String = "Select * From (Select Distinct CommonNos.PhoneNumber " + SelectQueryRefine + " From CommonNos " + FromQueryRefine + ") As tblCommons Where " + WhereQueryRefine
        Dim cmdRefineCommons As SqlCommand
        Dim DARefineCommons As New SqlDataAdapter
        Dim readerRefineCommons As SqlDataReader
        cmdRefineCommons = New SqlCommand(RefineCommonNosQuery, tempdbConn)
        'readerRefineCommons = cmdRefineCommons.ExecuteReader()
       
        'Try
        '    queryString1 = "Drop Table IF EXISTS [tbl_Intersection]"
        '    CreateTbCommand = New SqlCommand(queryString1, tempdbConn)
        '    CreateTbCommand.ExecuteNonQuery()
        'Catch ex03 As Exception
        'End Try

        'tblCommonNumbers = "CREATE TABLE [tbl_Intersection] (PhoneNumber nvarchar(255) null)"
        'OthersCommand = New SqlCommand(tblCommonNumbers, tempdbConn)
        'Try
        '    OthersCommand.ExecuteNonQuery()
        '    'TargetConnection.Close()
        'Catch ex02 As Exception
        '    MsgBox("Error creating Common Numbers table", MsgBoxStyle.OkOnly)
        '    'TargetConnection.Close()
        'End Try
        ''OthersCommand = New SqlCommand("Delete from [tbl_Intersection]", tempdbConn)
        ''OthersCommand.ExecuteNonQuery()
        'Dim intersectQueryCol1 As String = Nothing
        'Dim NextColQuery As String = Nothing
        'Dim bulkyCopy As SqlBulkCopy
        ''bulkyCopy.BatchSize = 50000

        'Dim CommonQuery As String = Nothing
        'Dim SelectQuery As String = "Select [tbl_Intersection].PhoneNumber"
        'Dim FromQuery As String = " From [tbl_Intersection] "
        'Dim JoinQuery As String = Nothing
        Dim cnicCommonNumbers As String = "Select CNIC, [CommonNumbers].* "
        'TotalNumberOfCDRs = DtColName.Rows.Count
        'TransferedFiles = 0
        'frm_Spy_Tech.prgbar_Common_Links.Maximum = TotalNumberOfCDRs
        'frm_Spy_Tech.prgbar_Common_Links.Value = 0
        'frm_Spy_Tech.lbCommonNos.Text = "Extracting common numbers " & TransferedFiles & " of " & TotalNumberOfCDRs
        'frm_Spy_Tech.Refresh()
        'If DtColName.Rows.Count > 0 Then
        '    For i As Integer = 0 To DtColName.Rows.Count - 1
        '        TransferedFiles = i + 1
        '        intersectQueryCol1 = "Select PhoneNumber from [" + DtColName.Rows(i)(0) + "] Intersect "
        '        For j As Integer = i + 1 To DtColName.Rows.Count - 1
        '            NextColQuery = intersectQueryCol1 + "Select PhoneNumber from [" + DtColName.Rows(j)(0) + "]"
        '            OthersCommand = New SqlCommand(NextColQuery, tempdbConn)
        '            OthersReader = OthersCommand.ExecuteReader()
        '            bulkyCopy = New SqlBulkCopy(TargetConnection)
        '            bulkyCopy.BatchSize = 50000
        '            bulkyCopy.BulkCopyTimeout = 0
        '            bulkyCopy.DestinationTableName = "[tbl_Intersection]"
        '            bulkyCopy.WriteToServer(OthersReader)
        '            OthersReader.Close()
        '        Next
        '        frm_Spy_Tech.prgbar_Common_Links.Value = frm_Spy_Tech.prgbar_Common_Links.Value + 1
        '        'frm_Spy_Tech.prgbar_Common_Links.Refresh()
        '        frm_Spy_Tech.lbCommonNos.Text = "Extracting common numbers " & TransferedFiles & " of " & TotalNumberOfCDRs
        '        frm_Spy_Tech.Refresh()
        '        SelectQuery = SelectQuery + ", [" + DtColName.Rows(i)(0) + "]"
        '        cnicCommonNumbers = cnicCommonNumbers + ", [" + DtColName.Rows(i)(0) + "]"
        '        JoinQuery = JoinQuery + " left join (Select [tbl_Intersection].PhoneNumber,COUNT([" + DtColName.Rows(i)(0) + "].PhoneNumber) as " + _
        '                    "[" + DtColName.Rows(i)(0) + "] from [tbl_Intersection] " + _
        '                    " left join [" + DtColName.Rows(i)(0) + "] on [" + DtColName.Rows(i)(0) + "].PhoneNumber=[tbl_Intersection].PhoneNumber " + _
        '                    " group by [tbl_Intersection].PhoneNumber) [" + DtColName.Rows(i)(0) + "] " + _
        '                    " on [" + DtColName.Rows(i)(0) + "].PhoneNumber=[tbl_Intersection].PhoneNumber"
        '    Next
        '    CommonQuery = SelectQuery + FromQuery + JoinQuery
        'End If
        ''Counting common numbers
        'Dim Da As New SqlDataAdapter
        'OthersCommand = New SqlCommand(CommonQuery, tempdbConn)
        ''OthersReader = OthersCommand.ExecuteReader()
        ''OthersCommand.ExecuteNonQuery()

        'Da.SelectCommand = OthersCommand
        ' readerRefineCommons = cmdRefineCommons.ExecuteReader()
        DARefineCommons.SelectCommand = cmdRefineCommons
        Dim dt As New System.Data.DataTable("tblCommonNumbers")
        Dim dtCNIC As New System.Data.DataTable
        'Dim Column1 As DataColumn = New DataColumn(TempTableName & "_")
        'Column1.DataType = System.Type.GetType("System.String")
        dt.Clear()
        DARefineCommons.Fill(dt)
        If frm_Spy_Tech.chbCNIC.Checked = True Then
            Dim BCopy As New SqlBulkCopy(db2021Connection)
            Dim daCNIC As New SqlDataAdapter
            'TargetConnection.Open()
            BCopy.BatchSize = 100000
            BCopy.BulkCopyTimeout = 0
            BCopy.DestinationTableName = "[CommonNumbers]"
            BCopy.WriteToServer(dt)
            ' Dim totalNumbers As Integer = dt.Rows.Count
            'ds.Tables.Add(dt)
            cnicCommonNumbers = cnicCommonNumbers + " From [CommonNumbers] left join Masterdb2023 on Masterdb2023.MSISDN=[CommonNumbers].CommonNumbers"
            OthersCommand = New SqlCommand(cnicCommonNumbers, db2021Connection)

            daCNIC.SelectCommand = OthersCommand
            dtCNIC.Clear()
            daCNIC.Fill(dtCNIC)
        End If
        For k As Integer = 0 To TotalNumberOfCDRs - 1
            OnlyFileNames(k) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(k))

            TempTableName = "_" & OnlyFileNames(k)
            TempTableName = TempTableName
            DtColName.Rows.Add(TempTableName)
            Try
                queryString1 = "Drop Table IF EXISTS [" & TempTableName & "]"
                CreateTbCommand = New SqlCommand(queryString1, tempdbConn)
                CreateTbCommand.ExecuteNonQuery()
            Catch ex04 As Exception
            End Try
        Next
        Try
            queryString1 = "Drop Table IF EXISTS [CommonNumbers]"
            CreateTbCommand = New SqlCommand(queryString1, tempdbConn)
            CreateTbCommand.ExecuteNonQuery()
        Catch ex05 As Exception
        End Try
        'OthersCommand = New SqlCommand("Delete from [tbl_Intersection]", tempdbConn)
        'OthersCommand.ExecuteNonQuery()
        If frm_Spy_Tech.chbCNIC.Checked = False Then
            If dt.Rows.Count = 0 Then
                MsgBox("Files do not have common numbers", vbOKOnly)
                frm_Spy_Tech.prgbar_Common_Links.Visible = False
                frm_Spy_Tech.lbCommonNos.Text = ""
                frm_Spy_Tech.lbCommonNos.Visible = False
                GC.Collect()
                Exit Function
            End If
            DatatableToExcel(dt, onlyXLFilePath)
            'BulkDataToExcel(dt, onlyXLFilePath)
        ElseIf frm_Spy_Tech.chbCNIC.Checked = True Then
            If dtCNIC.Rows.Count = 0 Then
                MsgBox("Files do not have common numbers", vbOKOnly)
                frm_Spy_Tech.prgbar_Common_Links.Visible = False
                frm_Spy_Tech.lbCommonNos.Text = ""
                frm_Spy_Tech.lbCommonNos.Visible = False
                GC.Collect()
                Exit Function
            End If
            DatatableToExcel_CNIC(dtCNIC, onlyXLFilePath)
            'BulkDataToExcel(dt, onlyXLFilePath)
        End If

    End Function
    Function releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Function
    Function DatatableToExcel(ByVal dtTemp As Data.DataTable, ByVal xlFilePath As String)
        Dim _excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        wBook = _excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        Dim dt As System.Data.DataTable = dtTemp
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 1
        Dim rowIndex As Integer = 1
        Dim colName As String
        Dim NoOfRows As Integer = dt.Rows.Count
        Dim InsertedRows As Integer = 0
        frm_Spy_Tech.prgbar_Common_Links.Maximum = NoOfRows + 1
        frm_Spy_Tech.prgbar_Common_Links.Value = 0
        frm_Spy_Tech.lbCommonNos.Text = "Transfered Records " + InsertedRows.ToString + " from " + NoOfRows.ToString
        _excel.Cells(1, colIndex) = "Sr.No."
        colName = ConvertToLetter(colIndex)
        _excel.Range(colName + "1 : " + colName + "1").Font.Bold = True
        '_excel.Range(colName + "1 : " + colName + "1").Interior.Color = System.Drawing.Color.DarkGray
        _excel.Range(colName + "1 : " + colName + "1").Borders.LineStyle = XlLineStyle.xlContinuous
        _excel.Range(colName + "1 : " + colName + "1").Borders.Weight = 3.0
        _excel.Cells.Range(colName & ":" & colName).NumberFormat = "#"
        _excel.Cells.Range(colName & ":" & colName).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        _excel.Cells.Range(colName & ":" & colName).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        For Each dc In dt.Columns
            colIndex = colIndex + 1
            If colIndex > 2 Then
                _excel.Cells(1, colIndex) = (dc.ColumnName).Substring(1)
                colName = ConvertToLetter(colIndex)

            ElseIf colIndex = 2 Then
                _excel.Cells(1, colIndex) = "Phone" + vbCrLf + "Number"
                colName = ConvertToLetter(colIndex)
            End If
            '_excel.Range(colName + "1 : " + colName + "1").Font.Color = System.Drawing.Color.White
            _excel.Range(colName + "1 : " + colName + "1").Font.Bold = True
            '_excel.Range(colName + "1 : " + colName + "1").Interior.Color = System.Drawing.Color.DarkGray
            _excel.Range(colName + "1 : " + colName + "1").Borders.LineStyle = XlLineStyle.xlContinuous
            _excel.Range(colName + "1 : " + colName + "1").Borders.Weight = 3.0
            _excel.Cells.Range(colName & ":" & colName).NumberFormat = "#"
            _excel.Cells.Range(colName & ":" & colName).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            _excel.Cells.Range(colName & ":" & colName).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        Next

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 1
            _excel.Cells(rowIndex, colIndex) = rowIndex - 1
            _excel.Cells(rowIndex, colIndex).Borders.LineStyle = XlLineStyle.xlContinuous
            _excel.Cells(rowIndex, colIndex).Borders.Weight = 2.0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                If dr(dc.ColumnName) <> 0 Then
                    _excel.Cells(rowIndex, colIndex) = dr(dc.ColumnName).ToString
                End If
                _excel.Cells(rowIndex, colIndex).Borders.LineStyle = XlLineStyle.xlContinuous
                _excel.Cells(rowIndex, colIndex).Borders.Weight = 2.0
            Next
            If rowIndex Mod 2 = 0 Then
                _excel.Range(ConvertToLetter(1) + rowIndex.ToString + " : " + ConvertToLetter(colIndex) + rowIndex.ToString).Interior.Color = System.Drawing.Color.LightGray
            End If
            InsertedRows = InsertedRows + 1
            frm_Spy_Tech.prgbar_Common_Links.Value = rowIndex
            frm_Spy_Tech.prgbar_Common_Links.Refresh()
            frm_Spy_Tech.lbCommonNos.Text = "Transfered Records " + InsertedRows.ToString + " from " + NoOfRows.ToString
            frm_Spy_Tech.lbCommonNos.Refresh()
        Next

        wSheet.Columns.AutoFit()
        _excel.Range("A1").Rows.EntireRow.Insert()
        _excel.Range("A1").Rows.EntireRow.Insert()
        Dim MergeRange As Range = _excel.Range(ConvertToLetter(1) + "1 : " + ConvertToLetter((dt.Columns.Count) + 1) + "1")
        '_excel.Range(colName(1) + "1 : " + colName(dt.Columns.Count + 1) + "1").Merge()
        MergeRange.Merge()
        MergeRange = _excel.Range(ConvertToLetter(1) + "1 : " + ConvertToLetter((dt.Columns.Count) + 1) + "1")
        MergeRange.FormulaR1C1 = "Common Numbers Report"
        MergeRange.Font.Size = 20
        MergeRange.Font.Bold = True
        MergeRange.RowHeight = 30
        MergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        MergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        MergeRange = _excel.Range(ConvertToLetter(1) + "2 : " + ConvertToLetter((dt.Columns.Count) + 1) + "2")
        MergeRange.Merge()
        MergeRange = _excel.Range(ConvertToLetter(1) + "2 : " + ConvertToLetter((dt.Columns.Count) + 1) + "2")
        MergeRange.NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
        MergeRange.FormulaR1C1 = Date.Now.ToString
        MergeRange.Font.Size = 14
        MergeRange.Font.Bold = True
        MergeRange.RowHeight = 25
        MergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        MergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

        MergeRange = _excel.Range(ConvertToLetter(1) + "3 : " + ConvertToLetter((dt.Columns.Count) + 1) + "3")
        MergeRange.Select()
        With _excel.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 3
        End With
        _excel.ActiveWindow.FreezePanes = True
        Dim strFileName As String = xlFilePath + "\CommonNumbersReport.xlsx"
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If

        wBook.SaveAs(strFileName)
        wBook.Close()
        _excel.Quit()
        MsgBox("File created " + strFileName, MsgBoxStyle.OkOnly)
        frm_Spy_Tech.prgbar_Common_Links.Visible = False
        frm_Spy_Tech.lbCommonNos.Text = ""
        frm_Spy_Tech.lbCommonNos.Visible = False
        GC.Collect()
    End Function

    Function ConvertToLetter(ByVal iCol As Integer) As String
        Dim col As Integer = iCol - 1
        Dim Reminder_Part As Integer = col Mod 26
        Dim Integer_Part As Integer = Int(col / 26)

        If Integer_Part = 0 Then
            ConvertToLetter = Chr(Reminder_Part + 65)
        Else
            ConvertToLetter = Chr(Integer_Part + 64) + Chr(Reminder_Part + 65)
        End If


    End Function
    Dim cnicReader As SqlDataReader
    Dim CNICQuery As String
    Function cnicToCommonNos(ByVal phnNumber As String) As String
        Dim CNIC As String = Nothing
        SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;") '("Server=" + ServerName + ";Database=MasterDB2022;User Id=sajjad;Password=rajpoot;")
        SubscriberVerifiedConnection.Open()
        CNICQuery = "Select CNIC from Masterdb2023 where MSISDN= '" + phnNumber + "'"
        SubscriberCNICCommand = New SqlCommand(CNICQuery, SubscriberVerifiedConnection)
        Try
            cnicReader = SubscriberCNICCommand.ExecuteReader
        Catch ex As Exception

        End Try

        While cnicReader.Read
            CNIC = cnicReader(0).ToString
        End While

        SubscriberVerifiedConnection.Close()
        cnicReader.Close()
        Return CNIC
    End Function

    Function DatatableToExcel_CNIC(ByVal dtTemp As Data.DataTable, ByVal xlFilePath As String)

        Dim _excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = _excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        Dim dt As System.Data.DataTable = dtTemp
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 1
        Dim rowIndex As Integer = 1
        Dim colName As String
        Dim NoOfRows As Integer = dt.Rows.Count
        Dim NoOfCols As Integer = dt.Columns.Count + 1
        Dim InsertedRows As Integer = 0
        frm_Spy_Tech.prgbar_Common_Links.Maximum = NoOfRows + 1
        frm_Spy_Tech.prgbar_Common_Links.Value = 0
        frm_Spy_Tech.lbCommonNos.Text = "Transfered Records " + InsertedRows.ToString + " from " + NoOfRows.ToString
        _excel.Cells(1, colIndex) = "Sr.No."
        colName = ConvertToLetter(colIndex)
        _excel.Range(colName + "1 : " + colName + "1").Font.Bold = True
        '_excel.Range(colName + "1 : " + colName + "1").Interior.Color = System.Drawing.Color.DarkGray
        _excel.Range(colName + "1 : " + colName + "1").Borders.LineStyle = XlLineStyle.xlContinuous
        _excel.Range(colName + "1 : " + colName + "1").Borders.Weight = 3.0
        _excel.Cells.Range(colName & ":" & colName).NumberFormat = "#"
        _excel.Cells.Range(colName & ":" & colName).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        _excel.Cells.Range(colName & ":" & colName).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        'colIndex = colIndex + 1
        '_excel.Cells(1, colIndex) = "CNIC"
        'colName = ConvertToLetter(colIndex)
        '_excel.Range(colName + "1 : " + colName + "1").Font.Bold = True
        ''_excel.Range(colName + "1 : " + colName + "1").Interior.Color = System.Drawing.Color.DarkGray
        '_excel.Range(colName + "1 : " + colName + "1").Borders.LineStyle = XlLineStyle.xlContinuous
        '_excel.Range(colName + "1 : " + colName + "1").Borders.Weight = 3.0
        '_excel.Cells.Range(colName & ":" & colName).NumberFormat = "#"
        '_excel.Cells.Range(colName & ":" & colName).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        '_excel.Cells.Range(colName & ":" & colName).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        For Each dc In dt.Columns
            colIndex = colIndex + 1
            If colIndex > 3 Then
                _excel.Cells(1, colIndex) = (dc.ColumnName).Substring(1)
                colName = ConvertToLetter(colIndex)
            ElseIf colIndex = 2 Then
                _excel.Cells(1, colIndex) = (dc.ColumnName)
                colName = ConvertToLetter(colIndex)
            ElseIf colIndex = 3 Then
                _excel.Cells(1, colIndex) = "Phone" + vbCrLf + "Number"
                colName = ConvertToLetter(colIndex)
            End If
            '_excel.Range(colName + "1 : " + colName + "1").Font.Color = System.Drawing.Color.White
            _excel.Range(colName + "1 : " + colName + "1").Font.Bold = True
            '_excel.Range(colName + "1 : " + colName + "1").Interior.Color = System.Drawing.Color.DarkGray
            _excel.Range(colName + "1 : " + colName + "1").Borders.LineStyle = XlLineStyle.xlContinuous
            _excel.Range(colName + "1 : " + colName + "1").Borders.Weight = 3.0
            _excel.Cells.Range(colName & ":" & colName).NumberFormat = "#"
            _excel.Cells.Range(colName & ":" & colName).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            _excel.Cells.Range(colName & ":" & colName).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        Next
        ' Dim cnicOfPhnNo As String
        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 1
            _excel.Cells(rowIndex, colIndex) = rowIndex - 1
            _excel.Cells(rowIndex, colIndex).Borders.LineStyle = XlLineStyle.xlContinuous
            _excel.Cells(rowIndex, colIndex).Borders.Weight = 2.0
            For Each dc In dt.Columns
                colIndex = colIndex + 1
                'If dr(dc.ColumnName) <> 0 Then
                '    If dc.ColumnName <> "PhoneNumber" Then
                '        'colIndex = colIndex + 1
                '        _excel.Cells(rowIndex, colIndex) = dr(dc.ColumnName).ToString
                '    Else
                '        cnicOfPhnNo = cnicToCommonNos(dr(dc.ColumnName).ToString)
                '        If cnicOfPhnNo <> Nothing Then
                '            _excel.Cells(rowIndex, colIndex) = cnicOfPhnNo
                '            _excel.Cells(rowIndex, colIndex).Borders.LineStyle = XlLineStyle.xlContinuous
                '            _excel.Cells(rowIndex, colIndex).Borders.Weight = 2.0
                '        End If
                '        colIndex = colIndex + 1
                _excel.Cells(rowIndex, colIndex) = dr(dc.ColumnName).ToString

                'End If
                ' End If

                _excel.Cells(rowIndex, colIndex).Borders.LineStyle = XlLineStyle.xlContinuous
                _excel.Cells(rowIndex, colIndex).Borders.Weight = 2.0
            Next
            If rowIndex Mod 2 = 0 Then
                _excel.Range(ConvertToLetter(1) + rowIndex.ToString + " : " + ConvertToLetter(colIndex) + rowIndex.ToString).Interior.Color = System.Drawing.Color.LightGray
            End If
            InsertedRows = InsertedRows + 1
            frm_Spy_Tech.prgbar_Common_Links.Value = rowIndex
            frm_Spy_Tech.lbCommonNos.Text = "Transfered Records " + InsertedRows.ToString + " from " + NoOfRows.ToString
            frm_Spy_Tech.Refresh()
        Next

        wSheet.Columns.AutoFit()
        _excel.Range("A1").Rows.EntireRow.Insert()
        _excel.Range("A1").Rows.EntireRow.Insert()
        Dim MergeRange As Range = _excel.Range(ConvertToLetter(1) + "1 : " + ConvertToLetter((dt.Columns.Count) + 1) + "1")
        '_excel.Range(colName(1) + "1 : " + colName(dt.Columns.Count + 1) + "1").Merge()
        MergeRange.Merge()
        MergeRange = _excel.Range(ConvertToLetter(1) + "1 : " + ConvertToLetter((dt.Columns.Count) + 1) + "1")
        MergeRange.FormulaR1C1 = "Common Numbers Report"
        MergeRange.Font.Size = 20
        MergeRange.Font.Bold = True
        MergeRange.RowHeight = 30
        MergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        MergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        MergeRange = _excel.Range(ConvertToLetter(1) + "2 : " + ConvertToLetter((dt.Columns.Count) + 1) + "2")
        MergeRange.Merge()
        MergeRange = _excel.Range(ConvertToLetter(1) + "2 : " + ConvertToLetter((dt.Columns.Count) + 1) + "2")
        MergeRange.NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
        MergeRange.FormulaR1C1 = Date.Now.ToString
        MergeRange.Font.Size = 14
        MergeRange.Font.Bold = True
        MergeRange.RowHeight = 25
        MergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        MergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

        MergeRange = _excel.Range(ConvertToLetter(1) + "3 : " + ConvertToLetter((dt.Columns.Count) + 1) + "3")
        MergeRange.Select()
        With _excel.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 3
        End With
        _excel.ActiveWindow.FreezePanes = True
        Dim strFileName As String = xlFilePath + "\CommonNumbersReport.xlsx"
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If

        wBook.SaveAs(strFileName)
        wBook.Close()
        _excel.Quit()

        MsgBox("File created " + strFileName, MsgBoxStyle.OkOnly)
        frm_Spy_Tech.prgbar_Common_Links.Visible = False
        frm_Spy_Tech.lbCommonNos.Text = ""
        frm_Spy_Tech.lbCommonNos.Visible = False
        GC.Collect()
    End Function
    Function BulkDataToExcel(ByVal dtTemp As Data.DataTable, ByVal xlFilePath As String)
        Dim _excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim osheets As Microsoft.Office.Interop.Excel.Worksheets
        wBook = _excel.Workbooks.Add(dtTemp)
        'wSheet = wBook.ActiveSheet("Sheet1")
        'wSheet = wBook.Worksheets("Sheet1")
        'osheets = wBook.Worksheets.Add(dtTemp, "Sheet1")
    End Function
End Module
