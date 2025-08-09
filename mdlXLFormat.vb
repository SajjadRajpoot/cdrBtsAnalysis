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
Module mdlXLFormat
    Sub EditCDR(ByVal Files() As String, ByVal Company As String)

    End Sub
    Function XLFormat(ByVal All_Files() As String)
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
        'For j As Integer = 0 To TotalNumberOfCDRs - 1

        OnlyFileNames(0) = System.IO.Path.GetFileNameWithoutExtension(All_Files(0))
        TempTableName = "_" & OnlyFileNames(0)

        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & All_Files(0) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        'populate table with data with 92
        Dim queryString As String = "SELECT * from [" & SheetName & "]"

        InsertCommand = New OleDb.OleDbCommand(queryString, o)
        Dim Da As New OleDbDataAdapter
        Da.SelectCommand = InsertCommand
        Dim dt As New System.Data.DataTable(TempTableName)
        'Dim Column1 As DataColumn = New DataColumn(TempTableName & "_")
        'Column1.DataType = System.Type.GetType("System.String")
        dt.Clear()
        Da.Fill(dt)
        frmXLFormat.dgvXLFormat.DataSource = dt.AsEnumerable.Take(10).CopyToDataTable
        Dim ComboboxHeaderCell As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn

        'ComboboxHeaderCell.Visible = True
        ComboboxHeaderCell.Items.Add("a")
        ComboboxHeaderCell.Items.Add("b")
        ComboboxHeaderCell.Items.Add("b1")
        ComboboxHeaderCell.Items.Add("b2")
        For i As Integer = 0 To frmXLFormat.dgvXLFormat.Columns.Count - 1
            frmXLFormat.dgvXLFormat.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            'frmXLFormat.dgvXLFormat.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
            'frmXLFormat.dgvXLFormat.Columns(i).DefaultCellStyle.BackColor = Color.LightPink
            'frmXLFormat.dgvXLFormat.Columns(i).DefaultCellStyle.BackColor = Color.LightPink


            'ComboboxHeaderCell.Location = frmXLFormat.dgvXLFormat.GetCellDisplayRectangle(i, -1, True).Location
            'comboBoxHeaderCell.Size = frmXLFormat.dgvXLFormat.Columns[i].HeaderCell.Size
            'comboBoxHeaderCell1.Text = "Column1";
        Next
        frmXLFormat.dgvXLFormat.BackgroundColor = Color.LightPink
        'ds.Tables.Add(dt)
        'Next
    End Function

    Function colHeads(ByVal All_Files() As String)
        Dim dsColumnName As New DataSet
        Dim TotalNumberOfCDRs As Integer = All_Files.Length()
        Dim OnlyFileNames(TotalNumberOfCDRs) As String
        Dim TempTableName As String
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim SheetName As String = "cdr$"
        Dim InsertCommand As OleDb.OleDbCommand
        Dim ds As New DataSet
        Dim dtQuery As String
        'For j As Integer = 0 To TotalNumberOfCDRs - 1

        OnlyFileNames(0) = System.IO.Path.GetFileNameWithoutExtension(All_Files(0))
        TempTableName = "_" & OnlyFileNames(0)

        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & All_Files(0) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        'populate table with data with 92
        'Dim queryString As String = "SELECT * from [" & SheetName & "]"
        Dim strSql As String = "SELECT column_name FROM INFORMATION_SCHEMA.Columns where TABLE_NAME =  [" & SheetName & "] "
        'Dim queryString As String = "SELECT * from [" & SheetName & "]"

        InsertCommand = New OleDb.OleDbCommand(strSql, o)
        Dim Da As New OleDbDataAdapter
        Da.SelectCommand = InsertCommand
        Dim dt As New System.Data.DataTable(TempTableName)
        'Dim dAdapter As New SqlClient.SqlDataAdapter(strSql, o)
        Da.Fill(dsColumnName)
        'Dim Da As New OleDbDataAdapter
        'Da.SelectCommand = InsertCommand
        'Dim dt As New System.Data.DataTable(TempTableName)
        'Dim Column1 As DataColumn = New DataColumn(TempTableName & "_")
        'Column1.DataType = System.Type.GetType("System.String")
        'dt.Clear()
        'Da.Fill(dt)
        'Dim ColumnsName_String As String = ""
        'For i As Integer = 0 To dsColumnName.Tables(0).Rows.Count - 1
        '    ColumnsName_String = ColumnsName_String & "," & dsColumnName.Tables(0).Rows(i).Item(0)
        'Next
        'If ColumnsName_String.StartsWith(",") Then
        '    ColumnsName_String = ColumnsName_String.Remove(0, 1)
        'End If
        'Return ColumnsName_String
        'dt.Clear()
        'Da.Fill(dt)
        frmXLFormat.dgvXLFormat.DataSource = dsColumnName.Tables(0)
        Dim ComboboxHeaderCell As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn

        'ComboboxHeaderCell.Visible = True
        ComboboxHeaderCell.Items.Add("a")
        ComboboxHeaderCell.Items.Add("b")
        ComboboxHeaderCell.Items.Add("b1")
        ComboboxHeaderCell.Items.Add("b2")
        For i As Integer = 0 To frmXLFormat.dgvXLFormat.Columns.Count - 1
            frmXLFormat.dgvXLFormat.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            'frmXLFormat.dgvXLFormat.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
            'frmXLFormat.dgvXLFormat.Columns(i).DefaultCellStyle.BackColor = Color.LightPink
            'frmXLFormat.dgvXLFormat.Columns(i).DefaultCellStyle.BackColor = Color.LightPink


            'ComboboxHeaderCell.Location = frmXLFormat.dgvXLFormat.GetCellDisplayRectangle(i, -1, True).Location
            'comboBoxHeaderCell.Size = frmXLFormat.dgvXLFormat.Columns[i].HeaderCell.Size
            'comboBoxHeaderCell1.Text = "Column1";
        Next
        frmXLFormat.dgvXLFormat.BackgroundColor = Color.LightPink
    End Function
End Module
