'Imports Microsoft.Office.Interop
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


Public Class frm_Spy_Tech

    Public DataInGrid As Boolean = False

    Private Sub frm_Spy_Tech_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        LoginForm1.Close()
    End Sub
    Private Sub frm_Spy_Tech_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.DVG_SpyTech.DefaultCellStyle.SelectionBackColor = Me.DVG_SpyTech.DefaultCellStyle.BackColor
        Me.DVG_SpyTech.DefaultCellStyle.SelectionForeColor = Me.DVG_SpyTech.DefaultCellStyle.ForeColor
        Me.Panel3.Dock = DockStyle.Fill
        ttSpyTech.SetToolTip(btnMergeExcelFiles, "BTS merger")
        ttSpyTech.SetToolTip(btnVerisys, "Save Verisys")
        ttSpyTech.SetToolTip(btnBTS, "BTS Analyzer")
        ttSpyTech.SetToolTip(btn_Save_in_MSWord, "Save as Word Doc")
        ttSpyTech.SetToolTip(btn_Search_CNC, "Search by CNC")
        ttSpyTech.SetToolTip(btn_Search_Phone_Number, "Search by Phone Number")
        ' Me.DVG_SpyTech.Dock = DockStyle.Fill
        'Me.DVG_SpyTech.Anchor = AnchorStyles.Right
        ' Lstbx_CNIC_PhoneNumbers.Height = Me.Bottom - Lstbx_CNIC_PhoneNumbers.Top
        'Dim myTop As Integer = Me.Top
        'Dim myHeight As Integer = Me.Height
        'Dim myBottom As Integer = Me.Bottom
        'Dim LstTop As Integer = Lstbx_CNIC_PhoneNumbers.Top
        'Dim LstHeight As Integer = Lstbx_CNIC_PhoneNumbers.Height
        'Dim ListBottom As Integer = Lstbx_CNIC_PhoneNumbers.Bottom

        Call CreateSearchStore()
        prgbarSearch.Left = Panel3.Width / 2 - prgbarSearch.Width / 2
        prgbarSearch.Top = Panel3.Height / 2 - prgbarSearch.Height / 2
        ' DVG_SpyTech.Left = Panel3.Width / 2 - DVG_SpyTech.Width / 2
        'DVG_SpyTech.Columns(0).Width = 100
        'DVG_SpyTech.Columns(1).Width = 300
        'DVG_SpyTech.Columns(2).Width = Panel3.Width - 350
        'DVG_SpyTech.Left = 0
        'DVG_SpyTech.Top = 0
        'DVG_SpyTech.Width = DVG_SpyTech.Columns(0).Width + DVG_SpyTech.Columns(1).Width + DVG_SpyTech.Columns(2).Width
        prgbarSearch.Visible = False
        DVG_SpyTech.Visible = False

        btn_Search_Phone_Number.Focus()
        'LoginForm1.Show()
    End Sub



    Private Sub btn_Search_Phone_Number_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_Phone_Number.Click
        'Me.Panel3.Dock = DockStyle.Fill
        ' DVG_SpyTech.Visible = False
        Lstbx_CNIC_PhoneNumbers.Items.Clear()
        If txt_Phone_Number.Text.Length < 9 Then

            MsgBox("Incomplete phone number", MsgBoxStyle.OkOnly)
            txt_Phone_Number.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        LockPhoneNumber = Nothing
        LockCNIC = Nothing
        IsCDR = False
        prgbarSearch.Minimum = 0
        prgbarSearch.Maximum = 41
        prgbarSearch.Value = 0
        LockCNIC = Nothing
        LockPhoneNumber = Nothing
        ' prgbarSearch.Visible = True
        If chk_Add_Result.Checked = True Then
            IsAddRecord = True
            'DT_PhoneNumber.Clear()
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Phone_Number.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        Else
            IsAddRecord = False
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            DVG_SpyTech.Visible = False
            'ExportIndoc.Close()
            Call CreateDocTitle("Analysis of: " & txt_Phone_Number.Text)
        End If
        DataInGrid = True
        If txt_Phone_Number.Text.Trim.StartsWith("0") Then
            txt_Phone_Number.Text = txt_Phone_Number.Text.Trim.Substring(1)
        End If
        ' Call PTCL_Search(txt_Phone_Number.Text)
        'If getTableName(txt_Phone_Number.Text.Trim) = "Masterdb2021" Then
        '    Call FindNumberDB2021(txt_Phone_Number.Text.Trim)
        '    'Call FindNumberDB2022(txt_Phone_Number.Text.Trim)
        'ElseIf getTableName(txt_Phone_Number.Text.Trim) = "Masterdb2022" Then
        'Call FindNumberDB2221(txt_Phone_Number.Text.Trim)
        'Else
        'Call FindNumberDB2022(txt_Phone_Number.Text.Trim)
        'End If
        'Call FindPhoneNumber(txt_Phone_Number.Text)
        Call FindNumberDB2021(txt_Phone_Number.Text.Trim)
        'following function is disabled for rasheed 
        Call FindFromOthers(, txt_Phone_Number.Text.Trim)
        prgbarSearch.Value = 41
        prgbarSearch.Visible = False
        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        ' DVG_SpyTech.Visible = True
        'Call DirectSaveDoc("E:\Example", txt_Phone_Number.Text)

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DVG_SpyTech_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DVG_SpyTech.CellFormatting
        Try
            If e.RowIndex > 0 And e.ColumnIndex = 0 Then
                If DVG_SpyTech.Item(0, e.RowIndex - 1).Value = e.Value Then
                    e.Value = ""

                ElseIf e.RowIndex < DVG_SpyTech.Rows.Count - 0 Then
                    'DVG_SpyTech.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.SkyBlue
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DVG_SpyTech_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DVG_SpyTech.CellPainting
        Try

        
        If e.ColumnIndex = 0 AndAlso e.RowIndex <> -1 Then
            Using gridBrush As Brush = New SolidBrush(Me.DVG_SpyTech.GridColor), backColorBrush As Brush = New SolidBrush(e.CellStyle.BackColor)
                Using gridLinePen As Pen = New Pen(gridBrush)
                    'clearing cell
                    e.Graphics.FillRectangle(backColorBrush, e.CellBounds)
                    'Drawing line of bottom border and right border of current cell
                    'if next row cell has different content only draw bottom border line of current cell
                    If e.RowIndex < DVG_SpyTech.Rows.Count - 2 AndAlso DVG_SpyTech.Rows(e.RowIndex + 1).Cells(0).Value.ToString <> e.Value.ToString() Then
                        e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
                    End If
                    'Drawing right border line of current cell
                    e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom)
                    'draw/fill content in current cell, and fill only one cell of multiple same cells
                    If Not e.Value Is Nothing Then
                        If e.RowIndex > 0 AndAlso DVG_SpyTech.Rows(e.RowIndex - 1).Cells(0).Value.ToString() = e.Value.ToString() Then
                        Else
                            e.Graphics.DrawString(CType(e.Value, String), e.CellStyle.Font, Brushes.Black, e.CellBounds.X + 2, e.CellBounds.Y + 5, StringFormat.GenericDefault)
                        End If
                    End If
                    e.Handled = True
                End Using
            End Using
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub


    Private Sub btn_Save_in_MSWord_Click(sender As System.Object, e As System.EventArgs) Handles btn_Save_in_MSWord.Click
        'If txt_Phone_Number.Text.Trim <> "" Then
        '    Call CreateDocTitle("Analysis of: " & txt_Phone_Number.Text.Trim)
        'ElseIf txt_CNC.Text.Trim <> "" Then
        '    Call CreateDocTitle("Analysis of: " & txt_CNC.Text.Trim)
        'End If
        If DVG_SpyTech.RowCount = 0 Then
            MsgBox("There is no record to save", MsgBoxStyle.OkOnly)
            Exit Sub

        End If

        Dim Path_FileName As String = Nothing
        Dim OnlyPath As String = Nothing
        Dim DocumentTitle As String = Nothing
        Dim FileName As String = Nothing
        Dim SaveFileDialog1 As New SaveFileDialog
        'If WOPFileName <> Nothing Then
        '    SaveFileDialog1.FileName = WOPFileName
        'End If
        SaveFileDialog1.OverwritePrompt = True

        SaveFileDialog1.Filter = "Word Document (*.docx)|*.docx"
        If txt_Phone_Number.Text <> "" Then
            SaveFileDialog1.FileName = txt_Phone_Number.Text
        Else
            SaveFileDialog1.FileName = txt_CNC.Text
        End If
        If SaveFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Path_FileName = SaveFileDialog1.FileName
            FileName = System.IO.Path.GetFileNameWithoutExtension(SaveFileDialog1.FileName)
            OnlyPath = System.IO.Path.GetDirectoryName(SaveFileDialog1.FileName)
            'DocumentTitle = DocumentTitle.Substring(0, DocumentTitle.Length - 4)
        Else
            Exit Sub
        End If

        Call DirectSaveDoc(OnlyPath, FileName)
        'Call exportAsWordDoc(DocumentTitle, Path_FileName, False)

        ' Call exportAsWordDoc(Path_FileName, False)




        'Dim oApp As Microsoft.Office.Interop.Word.Application
        'Dim oDoc As Document
        'Dim NumberOfChacracter As Integer
        'Dim title As String
        'Dim para1 As Paragraph
        'Dim para2 As Paragraph
        'oApp = CType(CreateObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
        'oDoc = oApp.Documents.Add()
        'oDoc.Range.Delete()
        'oDoc.Activate()
        'oDoc.PageSetup.LeftMargin = 36
        'oDoc.PageSetup.RightMargin = 36
        'Dim rng As Range = oDoc.Range(0, 0)
        ''Dim title As String = "Muhammad Sajjad Ahmad" & vbCrLf
        'title = "Analysis of :" & txt_Phone_Number.Text & vbCrLf & "Date:  " & Date.Today & "Time:  " & Format(Date.Now, "hh:mm:ss")
        'NumberOfChacracter = title.Length()
        ''rng.InsertAfter(title)

        ''rng.Font.Name = "Verdana"
        ''rng.Font.Size = 16
        ''rng.SetRange(0, 0)
        'para1 = oDoc.Content.Paragraphs.Add()
        'para1.Range.Text = title
        'para1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
        'para1.Range.Font.Size = 14
        'para1.Format.SpaceAfter = 12


        ''para2 = oDoc.Content.Paragraphs.Add()
        ''para2.Range.Text = "Date:  " & Date.Today & "Time:  " & Date.Now
        ''para2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        ''para2.Range.InsertParagraphAfter()
        ''rng = oDoc.Content
        'rng.SetRange(NumberOfChacracter + 1, NumberOfChacracter + 1)
        '' rng = oDoc.Content
        ''Dim i As Integer = rng.Paragraphs.Count()

        ''rng = oDoc.Content
        'Dim NoOfRows As Integer = DVG_SpyTech.RowCount + 1
        'Dim NoOfColumns As Integer = DVG_SpyTech.ColumnCount
        'Dim tbl As Table = oDoc.Tables.Add(Range:=rng, NumRows:=NoOfRows, NumColumns:=NoOfColumns)
        'NoOfRows = NoOfRows - 2
        'NoOfColumns = NoOfColumns - 1
        ''tbl.Range.Font.Name = "Arial"
        'tbl.Borders(WdBorderType.wdBorderBottom).Visible = True
        'tbl.Borders(WdBorderType.wdBorderHorizontal).Visible = True
        'tbl.Borders(WdBorderType.wdBorderLeft).Visible = True
        'tbl.Borders(WdBorderType.wdBorderRight).Visible = True
        'tbl.Borders(WdBorderType.wdBorderTop).Visible = True
        'tbl.Borders(WdBorderType.wdBorderVertical).Visible = True
        'tbl.Range.ParagraphFormat.LineUnitAfter = 0.01
        'tbl.Range.Font.Size = 10
        'tbl.Range.Font.Bold = True
        'tbl.Columns(1).Width = 80
        'tbl.Columns(2).Width = 320
        'tbl.Columns(3).AutoFit()
        'tbl.Rows(2).Alignment = WdRowAlignment.wdAlignRowLeft
        '' tbl.Cell(2, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        ''headers of table
        'tbl.Cell(1, 1).Range.Text = "Phone"
        'tbl.Cell(1, 1).Range.Font.Size = 11
        'tbl.Cell(1, 1).Range.Font.Bold = True
        'tbl.Cell(1, 2).Range.Text = "Detail"
        'tbl.Cell(1, 2).Range.Font.Size = 11
        'tbl.Cell(1, 2).Range.Font.Bold = True
        'tbl.Cell(1, 3).Range.Text = "Photograph"
        'tbl.Cell(1, 3).Range.Font.Size = 11
        'tbl.Cell(1, 3).Range.Font.Bold = True
        'Dim phnumber As String = DVG_SpyTech.Rows(0).Cells(1).Value()
        ''populate table with data
        'Dim InitialRow As Integer
        'Dim Initialization As Boolean = False
        'For r As Integer = 0 To NoOfRows
        '    For c As Integer = 0 To NoOfColumns
        '        If c = 2 Then
        '            If DVG_SpyTech.Rows(r).Cells(c).Value() <> vbNull Then
        '                tbl.Cell(r + 2, c + 1).Range.InlineShapes.AddPicture(DVG_SpyTech.Rows(r).Cells(c).Value())
        '            End If
        '        Else
        '            With tbl.Cell(r + 2, c + 1).Range
        '                If c = 0 And r > 0 Then
        '                    If DVG_SpyTech.Rows(r - 1).Cells(c).Value() = DVG_SpyTech.Rows(r).Cells(c).Value() Then
        '                        ' tbl.Cell(2, 1).Merge(tbl.Cell(3, 1))
        '                        'tbl.Cell(r + 1, c).Merge(tbl.Cell(r + 2, c))
        '                        If Initialization = False Then
        '                            Initialization = True
        '                            InitialRow = r - 1
        '                        End If
        '                        If r = NoOfRows Then
        '                            tbl.Cell(InitialRow + 2, c).Merge(tbl.Cell(r + 2, c))
        '                            Initialization = False
        '                        End If
        '                    Else
        '                        If Initialization = True Then
        '                            tbl.Cell(InitialRow + 2, c).Merge(tbl.Cell(r + 1, c))
        '                            Initialization = False
        '                        End If
        '                        .Text = DVG_SpyTech.Rows(r).Cells(c).Value()
        '                    End If
        '                Else
        '                    .Text = DVG_SpyTech.Rows(r).Cells(c).Value()
        '                    .Font.Size = 8
        '                End If


        '                If c = 1 Then
        '                    .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        '                End If
        '            End With


        '        End If


        '    Next
        'Next

        ''tbl.Style = "Table Professional"
        ''tbl.Rows(1)
        ''tbl.Rows(1).HeadingFormat = 0
        '' tbl.Cell(2, 2).Select()
        ''tbl.Cell(2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft


        ''tbl.Cell(2, 1).Range.Font.Size = 9


        ''tbl.Cell(2, 1).Range.Text = "03016258198"
        ''tbl.Cell(2, 2).Range.Text = "Village and post office Sheikh sukha"
        ''tbl.Cell(2, 3).Range.InlineShapes.AddPicture("E:\haider\Elite.jpg")
        'With oApp.Selection
        '    .SetRange(tbl.Range.End + 1, tbl.Range.End + 1)
        '    .Collapse(WdCollapseDirection.wdCollapseStart)
        'End With
        '' tbl.Cell(2, 1).Merge(tbl.Cell(3, 1))

        'oApp.Visible = True
    End Sub

    Private Sub txt_Phone_Number_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txt_Phone_Number.KeyDown
       
    End Sub

    Private Sub txt_Phone_Number_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Phone_Number.KeyPress
        'If IsNumeric(e.KeyChar) = True Or Asc(e.KeyChar) = vbBack Then
        '    Exit Sub
        'Else
        '    Beep()
        '    e.Handled = True

        'End If
        If e.KeyChar = vbCr Then
            btn_Search_Phone_Number_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Phone_Number_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txt_Phone_Number.KeyUp
        'If e.KeyCode = Keys.Enter Then
        '    e.SuppressKeyPress = True
        'End If
    End Sub

    Private Sub txt_Phone_Number_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Phone_Number.TextChanged

    End Sub

    Private Sub DVG_SpyTech_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DVG_SpyTech.CellContentClick

    End Sub

    Private Sub DVG_SpyTech_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles DVG_SpyTech.Paint
        
    End Sub

    Private Sub DVG_SpyTech_Resize(sender As Object, e As System.EventArgs) Handles DVG_SpyTech.Resize
        'Dim column1width As Integer = DVG_SpyTech.Columns(0).Width
        'Dim col2width As Integer = DVG_SpyTech.Columns(1).Width
        'Dim col3width As Integer = DVG_SpyTech.Columns(2).Width
        'DVG_SpyTech.Width = col2width + col3width + column1width + 100
        ''DVG_SpyTech.Left = (Panel3.Width - DVG_SpyTech.Width) / 2
    End Sub

    

    Private Sub btn_Search_CNC_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_CNC.Click

        If txt_CNC.Text.Length < 13 Then

            MsgBox("Please Enter the complete 13 digites of CNIC", MsgBoxStyle.OkOnly)
            txt_CNC.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        'Dim CNICRawal As String = txt_CNC.Text
        LockPhoneNumber = Nothing
        LockCNIC = Nothing
        IsCDR = False
        If chk_Add_Result.Checked = False Then
            IsAddRecord = False
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            DVG_SpyTech.Visible = False
            Call CreateDocTitle("Analysis of: " & txt_CNC.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Phone_Number.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
        Lstbx_CNIC_PhoneNumbers.Items.Clear()
        Lstbx_CNIC_PhoneNumbers.Items.Add("Numbers Against:" & txt_CNC.Text)
        Lstbx_CNIC_PhoneNumbers.Items.Add("")

        prgbarSearch.Minimum = 0
        prgbarSearch.Maximum = 50
        prgbarSearch.Value = 0
        Call FindCNICDB2020(txt_CNC.Text)
        ''Call FindCNIC(txt_CNC.Text)

        'following function is disabled for rasheed 
        Call FindFromOthers(txt_CNC.Text)
        If Lstbx_CNIC_PhoneNumbers.Items.Count > 1 Then

        End If
        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If

        Me.Cursor = Cursors.Default
    End Sub
    Dim PathAndFileName As String = Nothing
    Dim OnlyFileName As String = Nothing
    Public DocTitle As String
    Public PathForDoc As String
    Private Sub btn_Browse_CDRs_Click(sender As System.Object, e As System.EventArgs) Handles btn_Browse_CDRs.Click
        btn_Search_in_CDRs.Enabled = False
        Dim OpenFileDialog1 As New OpenFileDialog
        
        Dim FileExtention As String = Nothing
        OpenFileDialog1.Filter = "Excel Files |*.xlsx;*.xls;*.csv"
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
            txt_CDR_Report_Path.Text = PathAndFileName
            btn_Search_in_CDRs.Enabled = True
            'frmXLFormat.ShowDialog()
        Else
            Exit Sub
        End If
    End Sub
    Function ChangeFormat(ByVal FormatONOFF As Boolean)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Try


            xlWorkbook = xlWorkbooks.Open(txt_CDR_Report_Path.Text)
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "cdr" Then
                    xlSheet = xlWorkbook.Worksheets("cdr")
                    IsSheetRenamed = True
                    Exit For
                End If
            Next xlSheet
            If IsSheetRenamed = False Then
                MsgBox("Please rename the sheet as 'cdr' and browse again", MsgBoxStyle.OkOnly)
                GoTo Incomplete
            End If
            ' xlSheet = xlWorkbook.Worksheets("bts")
            Dim IsFormatChanged As Boolean = False
            Dim ColNumber As Integer = 1
            Dim ColName As String
            Dim ColRange As String
            While IsFormatChanged = False
                If xlSheet.Cells(1, ColNumber).Value = "b" Then
                    ColName = Chr(65 + ColNumber - 1)
                    ColRange = ColName & ":" & ColName
                    If FormatONOFF = False Then
                        xlSheet.Range(ColRange).NumberFormat = "@"
                    Else
                        xlSheet.Range(ColRange).NumberFormat = "###"
                    End If
                    IsFormatChanged = True
                End If
                ColNumber = ColNumber + 1
                If ColNumber > 8 Then
                    MsgBox("Please  rename as 'b' of time column and browse again", MsgBoxStyle.OkOnly)
                    GoTo Incomplete
                End If
            End While
            IsFormatChanged = False
            ColNumber = 1
            xlApp.DisplayAlerts = False
            While IsFormatChanged = False
                If xlSheet.Cells(1, ColNumber).Value = "Call Type" Then
                    ColName = Chr(65 + ColNumber - 1)
                    ColRange = ColName & ":" & ColName
                    xlSheet.Range(ColRange).NumberFormat = "@"
                    xlSheet.Range(ColRange).Replace("Incoming SMS", "Incoming", , , False)
                    xlSheet.Range(ColRange).Replace("Outgoing SMS", "Outgoing", , , False)
                    xlSheet.Range(ColRange).Replace("IncomingSMS", "InComing", , , False)
                    xlSheet.Range(ColRange).Replace("OutgoingSMS", "Outgoing", , , False)
                    IsFormatChanged = True
                End If
                ColNumber = ColNumber + 1
                If ColNumber > 15 Then
                    MsgBox("Please set then name 'Call Type' of A Party column and browse again", MsgBoxStyle.OkOnly)
                    GoTo Incomplete
                End If
            End While


Incomplete:
            xlApp.DisplayAlerts = False
            ' xlSheet.SaveAs(txt_BTS_Path.Text)
            ' xlApp.DisplayAlerts = True
            xlWorkbook.SaveAs(txt_CDR_Report_Path.Text, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
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
    'Public NumbersOfCalls As String = Nothing
    'Public IsCDR As Boolean = False
    'Public Network As String = Nothing
    ' Public DateOfActivation As String = Nothing
    Dim TargetConnection As SqlConnection
    'Dim SearchIMEIConnection As SqlConnection
    Dim QueryString As String
    Dim CreateTbCommand As SqlCommand

    Sub ConnectionOpen()
        TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
        If TargetConnection.State = ConnectionState.Closed Then
            TargetConnection.Open()
        End If
    End Sub

    Sub CreateTBL_CallSummary(ByVal tblName As String)
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
            QueryString = "CREATE TABLE " & tblName & "(b varchar(35), [Call Type] varchar(50))"
            'QueryString = "CREATE TABLE " & tblName & "(" & ColumnsHeaders & ")"
            CreateTbCommand = New SqlCommand(QueryString, TargetConnection)
            CreateTbCommand.ExecuteNonQuery()
            TargetConnection.Close()
            TargetConnection.Dispose()
            CreateTbCommand.Dispose()
        Catch ex As Exception

        End Try
    End Sub
    Dim InsertQuery As String
    Sub InsertToCallSummary(ByVal tblName As String, ByVal b As String, ByVal CallType As String)
        Call ConnectionOpen()
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO [" & tblName & "] VALUES ('" & b & "','" & CallType & "')"
        'InsertQuery = "INSERT INTO [" & tblName & "] VALUES (" & ColumnsValues & ")"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        Try
            InsertIMEICommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error in insert", MsgBoxStyle.Information)
        End Try

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Sub CloseFileOpened(ByVal FileAddress As String)
        If IO.File.Exists(FileAddress) Then
            Dim Proced As Boolean = False
            Dim XlApp As Excel.Application = Nothing
            Dim XlWorkbooks As Excel.Workbooks = Nothing
            Dim xlWorkbook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorksheets As Excel.Worksheets = Nothing
            XlApp = New Excel.Application
            XlApp.DisplayAlerts = False
            XlWorkbooks = XlApp.Workbooks
            xlWorkbook = XlWorkbooks.Open(FileAddress)
            XlApp.Visible = False

            xlWorkbook.Close()
            XlApp.Quit()

        End If
        
    End Sub
    Sub CallSammary(ByVal xlFileAddress As String)
        'lbProgress.Text = "Initializing..."
        'prgbar_CDR_report.BackColor = Color.Transparent
        prgbar_CDR_report.Minimum = 0
        prgbar_CDR_report.Maximum = 5
        prgbar_CDR_report.Value = 0
        prgbar_CDR_report.Visible = True
        'lbProgress.Visible = True
        'Call ChangeFormat(False)
        Dim CountCallTypeSqlcon As SqlConnection
        Dim cmdCallSummary As SqlCommand
        Dim ReaderCallSummary As SqlDataReader
        Dim Call_aReader As OleDb.OleDbDataReader
        Dim Call_acommand As OleDb.OleDbCommand
        Dim Call_aQuery As String

        Dim SaveFilePath As String = txt_CDR_Report_Path.Text
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Dim SheetName As String = Nothing
        Dim TableName As String = "tblCallSummary"
        prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1
        Try
            xlWorkbook = xlWorkbooks.Open(txt_CDR_Report_Path.Text)
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "bts" Then
                    SheetName = "bts$"
                    IsSheetRenamed = True
                    Exit For
                ElseIf xlSheet.Name = "cdr" Then
                    SheetName = "cdr$"
                    IsSheetRenamed = True
                    Exit For
                End If

            Next xlSheet

        Catch ex As Exception
            MsgBox("Please check the sheet name; sheet name must be 'cdr' or 'bts'", MsgBoxStyle.Information)
            xlWorkbook.Close()
            xlWorkbooks.Close()
            xlApp.Quit()
            Exit Sub
        End Try

        xlWorkbook.Close()
        xlWorkbooks.Close()
        xlApp.Quit()
        prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1
        Dim InGrandTotal, OutGrandToatal, InRecordTotal, OutRecordTotal, GrandTotal As Integer
        InGrandTotal = 0
        OutGrandToatal = 0
        InRecordTotal = 0
        OutRecordTotal = 0

        GrandTotal = 0
        Dim RowLabel As String = Nothing
        If IsSheetRenamed = False Then
            MsgBox("Please check the sheet name; sheet name must be 'cdr' or 'bts'", MsgBoxStyle.Information)
            Exit Sub
        End If
        DVG_SpyTech.Visible = False
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim ColumnName As String = "b"
        'Dim SheetName As String = "cdr$"
        Dim totalRows As Integer = 0
        Call CreateTBL_CallSummary(TableName)
        ' Dim queryString As String = "select [" & ColumnName & "] as party, count(*) as CountOf from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        ' Dim queryString As String = "select [" & ColumnName & "] as party, count(*) as CountOf from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        Dim CountqueryString As String '= "select COUNT(*) as CountOf from (select [" & ColumnName1 & "] as Party from [" & SheetName & "] GROUP BY [" & ColumnName1 & "]" + _
        '                "UNION ALL select [" & ColumnName2 & "] as Party from [" & SheetName & "] GROUP BY [" & ColumnName2 & "] )"
        'Dim queryString As String = "select [" & ColumnName & "]  from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        CountCallTypeSqlcon = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
        'If CountCallTypeSqlcon.State = ConnectionState.Closed Then
        '    TargetConnection.Open()
        'End If
        Dim ExcelToSQL_Query As String = "Select b , [Call Type] from [" & SheetName & "]"
        Call_aQuery = "select DISTINCT [" & ColumnName1 & "]  from [" & SheetName & "]" 'GROUP BY [" & ColumnName & "]"
        'Dim queryString As String = "select  DISTINCT [" & ColumnName & "]  from [" & SheetName & "]" 'GROUP BY [" & ColumnName & "]"
        Dim queryString As String = "select  DISTINCT [" & ColumnName & "]  from [" & TableName & "]" 'GROUP BY [" & ColumnName & "]"
        'CountqueryString = "select COUNT(*) AS CountOf From (SELECT DISTINCT [" & ColumnName & "] from [" & SheetName & "])" 'GROUP BY [" & ColumnName & "]"
        CountqueryString = "SELECT COUNT(DISTINCT [" & ColumnName & "]) AS MYCOUNT from [" & TableName & "]" 'GROUP BY [" & ColumnName & "]"
        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & txt_CDR_Report_Path.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        ' Dim ocmd As New OleDb.OleDbCommand(queryString, o)
        Dim ocmd As New OleDb.OleDbCommand(ExcelToSQL_Query, o)
        Dim ocmd1 As OleDb.OleDbCommand
        'ocmd1 = New OleDb.OleDbCommand(CountqueryString, o)

        'Dim oreader As OleDb.OleDbDataReader
        Dim oreader As SqlDataReader
        Dim cmdCoundCallType As SqlCommand
        Dim bParty As String
        Dim CallTypes As String
        'oreader = ocmd1.ExecuteReader
        'While oreader.Read
        '    prgbar_CDR_report.Maximum = oreader(0) + 10
        'End While
        'oreader.Close()
        'ocmd1.Dispose()
        prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1



        Dim oreader1 As OleDb.OleDbDataReader
        Dim recCount As Integer = 0
        Try
            oreader1 = ocmd.ExecuteReader()

            While oreader1.Read
                If IsDBNull(oreader1(0)) = False Then

                    bParty = Trim(oreader1(0)).ToString
                    If bParty.Length >= 10 Then
                        bParty = bParty.Substring(bParty.Length - 10, 10)
                    End If
                Else
                    bParty = ""
                End If
                If IsDBNull(oreader1(1)) = False Then
                    If Trim(oreader1(1).ToString).ToUpper = "INCOMING SMS" Or Trim(oreader1(1).ToString).ToUpper = "INCOMINGSMS" Or Trim(oreader1(1).ToString).ToUpper = "INCOMING" Or Trim(oreader1(1).ToString).ToUpper.Contains("INCOMING") Then
                        CallTypes = "INCOMING"
                    ElseIf Trim(oreader1(1).ToString).ToUpper = "OUTGOING SMS" Or Trim(oreader1(1).ToString).ToUpper = "OUTGOINGSMS" Or Trim(oreader1(1).ToString).ToUpper = "OUTGOING" Or Trim(oreader1(1).ToString).ToUpper.Contains("OUTGOING") Then
                        CallTypes = "OUTGOING"
                    End If

                Else
                    bParty = "UNKNOWN"
                End If
                Call InsertToCallSummary(TableName, bParty, CallTypes)
                'If IsDBNull(oreader1(0)) = False And IsDBNull(oreader1(1)) = False Then
                '    Call InsertToCallSummary(TableName, Trim(oreader1(0)).ToString, Trim(oreader1(1)).ToString)
                'ElseIf IsDBNull(oreader1(0)) = True And IsDBNull(oreader1(1)) = False Then
                '    Call InsertToCallSummary(TableName, "", Trim(oreader1(1)).ToString)
                'ElseIf IsDBNull(oreader1(0)) = False And IsDBNull(oreader1(1)) = True Then

                'End If
                ' Call InsertToCallSummary("tblCallSummary", Trim(oreader1(0)).ToString, Trim(oreader1(1)).ToString)
                recCount = recCount + 1
            End While
        Catch ex As Exception
            MsgBox("Error in creating Call summary " & Err.Description)
        End Try
        prgbar_CDR_report.Value = 5
        recCount = 0
        Call ConnectionOpen()
        cmdCallSummary = New SqlCommand(CountqueryString, TargetConnection)
        ReaderCallSummary = cmdCallSummary.ExecuteReader

        While ReaderCallSummary.Read
            If IsDBNull(ReaderCallSummary(0)) = False Then
                prgbar_CDR_report.Maximum = ReaderCallSummary(0) + 15
                'recCount = recCount + 1
            End If
        End While
        lbProgress.Text = "Creating Call Summary....."
        ' prgbar_CDR_report.BackColor = Color.Transparent
        prgbar_CDR_report.Minimum = 0

        prgbar_CDR_report.Value = 0
        TargetConnection.Close()
        ReaderCallSummary.Close()
        cmdCallSummary.Dispose()
        'prgbar_CDR_report.Maximum = recCount + 15
        Call CreateNewExcelFile(xlFileAddress)
        newXlWorkSheet.Cells(3, 1) = "Name"
        newXlWorkSheet.Cells(3, 2) = "Address"
        newXlWorkSheet.Cells(3, 3) = "CNIC"
        newXlWorkSheet.Cells(3, 4) = "b"
        newXlWorkSheet.Cells(3, 5) = "Incoming"
        newXlWorkSheet.Cells(3, 6) = "Outgoing"
        newXlWorkSheet.Cells(3, 7) = "GrandTotal"
        newXlWorkSheet.Range("A3:G3").Font.Size = 14
        newXlWorkSheet.Range("A3:G3").Interior.Color = Color.LightBlue

        For i As Integer = 1 To 6
            With newXlWorkSheet.Range(Convert.ToChar(64 + i) & 3).Borders(XlBordersIndex.xlEdgeRight)
                .LineStyle = XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlThin
            End With
        Next

        Call_acommand = New OleDb.OleDbCommand(Call_aQuery, o)
        Call_aReader = Call_acommand.ExecuteReader
        Dim aParty As String
        While Call_aReader.Read
            If Trim(Call_aReader(0).ToString).Length >= 10 And IsDBNull(Call_aReader(0)) = False Then
                aParty = Trim(Call_aReader(0).ToString)
                Exit While
            End If
        End While

        GprCNIC = Nothing
        Network = Nothing
        SubName = Nothing
        DateOfActivation = Nothing
        If GprCNIC = Nothing Then
            IsCNIC_Need = True
            'If getTableName(Trim(aParty)) = "Masterdb2021" Then
            '    Call FindNumberDB2021(Trim(aParty))
            '    'Call FindNumberDB2022(txt_Phone_Number.Text.Trim)
            'ElseIf getTableName(Trim(aParty)) = "Masterdb2022" Then
            '    Call FindNumberDB2221(Trim(aParty))
            'Else
            '    Call FindNumberDB2022(Trim(aParty))
            'End If
            Call FindNumberDB2021(Trim(aParty))
            '' Call FindPhoneNumber(Trim(aParty))
            IsCNIC_Need = False
        End If

        newXlWorkSheet.Range("A1:G1").Merge()
        newXlWorkSheet.Range("A1:G1").HorizontalAlignment = Excel.Constants.xlCenter
        newXlWorkSheet.Range("A1:G1").Font.Size = 16
        newXlWorkSheet.Range("A1:G1").Font.Bold = True
        newXlWorkSheet.Range("A1:G1").NumberFormat = "###"
        newXlWorkSheet.Range("A2:G2").Merge()
        newXlWorkSheet.Range("A2:G2").HorizontalAlignment = Excel.Constants.xlCenter
        newXlWorkSheet.Range("A2:G2").Font.Size = 8
        'newXlWorkSheet.Range("A2:G2").Font.Bold = True
        Dim TitleNumber As String '= Trim(aParty.ToString)
        Dim isSubTilte As Boolean = False
        If SubName <> Nothing Then
            'newXlWorkSheet.Cells(1, 1) = Network
            TitleNumber = "(" & SubName
            isSubTilte = True
        Else
            'newXlWorkSheet.Cells(1, 1) = "NA"
            isSubTilte = False
        End If
        If DateOfActivation <> Nothing Then
            If isSubTilte = True Then
                'newXlWorkSheet.Cells(1, 2) = DateOfActivation
                TitleNumber = TitleNumber & " , " & DateOfActivation
                isSubTilte = True
            Else
                TitleNumber = "(" & DateOfActivation
                isSubTilte = True
            End If
        Else
            'newXlWorkSheet.Cells(1, 2) = "NA"
            If isSubTilte = False Then
                isSubTilte = False
            End If

        End If
        If GprCNIC <> Nothing Then
            'newXlWorkSheet.Cells(1, 3) = GprCNIC
            If isSubTilte = True Then
                'newXlWorkSheet.Cells(1, 2) = DateOfActivation
                TitleNumber = TitleNumber & " , " & GprCNIC & ")"
                isSubTilte = True
            Else
                TitleNumber = "(" & GprCNIC & ")"
                isSubTilte = True
            End If
        Else
            'newXlWorkSheet.Cells(1, 3) = "NA"
            isSubTilte = False
        End If
        newXlWorkSheet.Cells(1, 1) = Trim(aParty.ToString)
        newXlWorkSheet.Cells(2, 1) = TitleNumber
        newXlWorkSheet.Range("C1:F1").NumberFormat = "###"
        Dim IndexOfRow As Integer = 4
        Dim IsCountSet As Boolean = False
        prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1
        Call ConnectionOpen()
        cmdCallSummary = New SqlCommand(queryString, TargetConnection)
        ReaderCallSummary = cmdCallSummary.ExecuteReader

        CountCallTypeSqlcon.Open()
        While ReaderCallSummary.Read
            If IsCountSet = False Then

                IsCountSet = True
            End If
            If IsDBNull(ReaderCallSummary(0)) = False Then 'And oreader1(0).ToString.Length >= 10 Then
                Try
                    CountqueryString = "select [Call Type]  , count(*) as CallCount from [" & TableName & "] where b = '" & ReaderCallSummary(0).ToString & "' GROUP BY [Call Type]"
                Catch ex As Exception
                    CountqueryString = "select [Call Type]  , count(*) as CallCount from [" & TableName & "] where b = " & ReaderCallSummary(0).ToString & " GROUP BY [Call Type]"
                End Try

                'ocmd1 = New OleDb.OleDbCommand(CountqueryString, o)
                cmdCoundCallType = New SqlCommand(CountqueryString, CountCallTypeSqlcon)
                Try
                    oreader = cmdCoundCallType.ExecuteReader
                Catch ex As Exception
                    Try
                        CountqueryString = "select [Call Type]  , count(*) as CallCount from [" & TableName & "] where b = '" & Str(ReaderCallSummary(0)).ToString & "' GROUP BY [Call Type]"
                    Catch ex1 As Exception
                        CountqueryString = "select [Call Type]  , count(*) as CallCount from [" & TableName & "] where b = '" & ReaderCallSummary(0).ToString & "' GROUP BY [Call Type]"
                    End Try
                    'CountqueryString = "select [Call Type]  , count(*) as CallCount from [" & SheetName & "] where b = '" & Str(Trim(oreader1(0))).ToString & "' GROUP BY [Call Type]"
                    'ocmd1 = New OleDb.OleDbCommand(CountqueryString, o)
                    cmdCoundCallType = New SqlCommand(CountqueryString, CountCallTypeSqlcon)
                    Try
                        oreader = cmdCoundCallType.ExecuteReader
                    Catch ex2 As Exception
                        MsgBox("Please make sure the name of Columns 'b' and 'Call Type' ", MsgBoxStyle.Information)
                        newXlApp.DisplayAlerts = False
                        newXlWorkbooks = Nothing
                        newXlApp.Quit()
                        newXlApp = Nothing
                        o.Close()
                        Exit Sub
                    End Try


                End Try

                InRecordTotal = 0
                OutRecordTotal = 0
                While oreader.Read
                    If Trim(oreader(0).ToString).Length >= 8 Then
                        If Trim(oreader(0).ToString.ToUpper).Contains("INCOMING") = True Then
                            InGrandTotal = InGrandTotal + oreader(1)
                            InRecordTotal = oreader(1)
                        ElseIf Trim(oreader(0).ToString.ToUpper).Contains("OUTGOING") = True Then
                            OutGrandToatal = OutGrandToatal + oreader(1)
                            OutRecordTotal = oreader(1)
                        End If
                    End If
                End While
                oreader.Close()
                cmdCoundCallType.Dispose()
                GprCNIC = Nothing
                Network = Nothing
                SubName = Nothing
                DateOfActivation = Nothing
                If GprCNIC = Nothing Then
                    IsCNIC_Need = True
                    If Trim(ReaderCallSummary(0)).ToString.Length >= 10 Then
                        'If getTableName(Trim(ReaderCallSummary(0).ToString)) = "Masterdb2021" Then
                        '    Call FindNumberDB2021(Trim(ReaderCallSummary(0).ToString))
                        '    'Call FindNumberDB2022(txt_Phone_Number.Text.Trim)
                        'ElseIf getTableName(Trim(ReaderCallSummary(0).ToString)) = "Masterdb2022" Then
                        '    Call FindNumberDB2221(Trim(ReaderCallSummary(0).ToString))
                        'Else
                        '    Call FindNumberDB2022(Trim(ReaderCallSummary(0).ToString))
                        'End If
                        Call FindNumberDB2021(Trim(ReaderCallSummary(0).ToString))
                        ''Call FindPhoneNumber(Trim(ReaderCallSummary(0).ToString))
                    End If
                    IsCNIC_Need = True
                End If
                If SubName <> Nothing Then
                    newXlWorkSheet.Cells(IndexOfRow, 1) = SubName
                End If
                If DateOfActivation <> Nothing Then
                    If DateOfActivation.StartsWith("=") Or DateOfActivation.StartsWith(".") Or DateOfActivation.StartsWith("?") Then

                        newXlWorkSheet.Cells(IndexOfRow, 2) = DateOfActivation.Remove(0)
                    Else
                        newXlWorkSheet.Cells(IndexOfRow, 2) = DateOfActivation
                    End If
                End If
                If GprCNIC <> Nothing Then
                    newXlWorkSheet.Cells(IndexOfRow, 3) = GprCNIC
                End If
                'If Trim(oreader1(0).ToString).Length >= 10 Then
                '    newXlWorkSheet.Cells(IndexOfRow, 4) = Trim(oreader1(0).ToString) '.Substring(Trim(oreader1(0).ToString).Length - 10, 10)
                'Else

                newXlWorkSheet.Cells(IndexOfRow, 4) = Trim(ReaderCallSummary(0).ToString)

                'End If
                If InRecordTotal > 0 Then
                    newXlWorkSheet.Cells(IndexOfRow, 5) = InRecordTotal
                End If
                If OutRecordTotal > 0 Then
                    newXlWorkSheet.Cells(IndexOfRow, 6) = OutRecordTotal
                End If
                If InRecordTotal + OutRecordTotal > 0 Then
                    newXlWorkSheet.Cells(IndexOfRow, 7) = InRecordTotal + OutRecordTotal
                End If

                With newXlWorkSheet.Range("A4:G" & IndexOfRow).Borders(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                For i As Integer = 1 To 6
                    With newXlWorkSheet.Range(Convert.ToChar(64 + i) & IndexOfRow).Borders(XlBordersIndex.xlEdgeRight)
                        .LineStyle = XlLineStyle.xlContinuous
                        .Weight = Excel.XlBorderWeight.xlThin
                    End With
                Next
                IndexOfRow = IndexOfRow + 1
            End If
            newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).HorizontalAlignment = Excel.Constants.xlCenter
            prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1
        End While
        oreader.Close()
        oreader1.Close()
        o.Close()
        ocmd.Dispose()
        cmdCoundCallType.Dispose()
        CountCallTypeSqlcon.Close()
        TargetConnection.Close()
        newXlWorkSheet.Cells(IndexOfRow, 4) = "Grand Total"
        If InGrandTotal > 0 Then
            newXlWorkSheet.Cells(IndexOfRow, 5) = "=sum(E4:E" & IndexOfRow - 1 & ")"
        End If
        If OutGrandToatal > 0 Then
            newXlWorkSheet.Cells(IndexOfRow, 6) = "=sum(F4:F" & IndexOfRow - 1 & ")"
        End If
        If InGrandTotal + OutGrandToatal > 0 Then
            newXlWorkSheet.Cells(IndexOfRow, 7) = "=sum(G4:G" & IndexOfRow - 1 & ")"
        End If
        newXlWorkSheet.Range("C4:F" & IndexOfRow).NumberFormat = "###"
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Font.Size = 14
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Interior.Color = Color.LightBlue
        newXlWorkSheet.Range("A3:G" & IndexOfRow).BorderAround(LineStyle:=Excel.XlLineStyle.xlContinuous, Weight:=Excel.XlBorderWeight.xlThick)
        'With newXlWorkSheet.Range("A4:G" & IndexOfRow - 1).Borders(XlBordersIndex.xlEdgeBottom)
        '    .LineStyle = XlLineStyle.xlContinuous
        '    .Weight = Excel.XlBorderWeight.xlThin
        'End With
        newXlWorkSheet.Range("A3:B3").HorizontalAlignment = Excel.Constants.xlCenter

        newXlWorkSheet.Range("A4:B" & IndexOfRow).HorizontalAlignment = Excel.Constants.xlLeft
        newXlWorkSheet.Range("A4:B" & IndexOfRow).Font.Size = 8
        newXlWorkSheet.Range("C3:G" & IndexOfRow).HorizontalAlignment = Excel.Constants.xlCenter

        newXlWorkSheet.Range("A3:G" & IndexOfRow).EntireColumn.AutoFit()

        For i As Integer = 1 To 6
            With newXlWorkSheet.Range(Convert.ToChar(64 + i) & IndexOfRow).Borders(XlBordersIndex.xlEdgeRight)
                .LineStyle = XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlThin
            End With
        Next
        IndexOfRow = IndexOfRow + 1
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Merge()
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).HorizontalAlignment = Excel.Constants.xlCenter
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Font.Size = 16
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Font.Bold = True
        newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).NumberFormat = "###"
        newXlWorkSheet.Cells(IndexOfRow, 1) = "IMEI"
        'newXlWorkSheet.Range("A2:G2").HorizontalAlignment = Excel.Constants.xlCenter
        'newXlWorkSheet.Range("A2:G2").Font.Size = 14
        ColumnName = "IMEI"
        queryString = "select  DISTINCT [" & ColumnName & "]  from [" & SheetName & "]"
        o.Open()
        ocmd = New OleDb.OleDbCommand(queryString, o)
        oreader1 = ocmd.ExecuteReader
        While oreader1.Read
            IndexOfRow = IndexOfRow + 1
            newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Merge()
            newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).HorizontalAlignment = Excel.Constants.xlCenter
            newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).Font.Size = 14
            newXlWorkSheet.Range("A" & IndexOfRow & ":G" & IndexOfRow).NumberFormat = "###"
            newXlWorkSheet.Cells(IndexOfRow, 1) = Trim(oreader1(0).ToString)
        End While
        oreader1.Close()
        ocmd.Dispose()
        o.Close()
        Dim counter As Integer = 0
        newXlApp.DisplayAlerts = False
        Dim Xl_file As String = xlFileAddress 'SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary")
        Try
            'Call CloseFileOpened(Xl_file)
            '    While File.Exists(Xl_file) = True
            '        counter = counter + 1
            '        Xl_file = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary" & counter.ToString)
            '    End While
            'newXlWorkbook.SaveAs(Xl_file, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            'newXlWorkSheet.SaveAs(Xl_file)
            'newXlWorkSheet.SaveAs(SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary"))
            newXlWorkbook.Save()
            prgbar_CDR_report.Value = prgbar_CDR_report.Maximum
            If chkBothFiles.Checked = False Then
                'MsgBox("File has been created: " & SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary"), MsgBoxStyle.Information)
                MsgBox("File has been created: " & Xl_file, MsgBoxStyle.Information)
                IsChkBothFiles = False
            Else
                'SaveMsg = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary")
                SaveMsg = Xl_file
                IsChkBothFiles = True
            End If
            newXlWorkbook.Close()
            newXlWorkbooks = Nothing
            newXlApp.Quit()
            newXlApp = Nothing
            releaseObject(newXlWorkSheet)
            releaseObject(newXlWorkbook)
            releaseObject(newXlWorkbooks)
            releaseObject(newXlApp)
            Exit Sub
            'MsgBox("File has been created: " & SaveFilePath, MsgBoxStyle.Information)
            prgbar_CDR_report.Value = prgbar_CDR_report.Maximum
        Catch ex As Exception
            'MsgBox("Please close the file " & SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary" & Err.Description), MsgBoxStyle.Information)
            MsgBox("Please close the file " & Xl_file & Err.Description, MsgBoxStyle.Information)
            'newXlWorkSheet.SaveAs(SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary"))
            newXlWorkbook.Save()
            prgbar_CDR_report.Value = prgbar_CDR_report.Maximum
            If chkBothFiles.Checked = False Then
                'MsgBox("File has been created: " & SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary"), MsgBoxStyle.Information)
                MsgBox("Please close the file " & Xl_file & Err.Description, MsgBoxStyle.Information)
                IsChkBothFiles = False
            Else
                'SaveMsg = SaveFilePath.Insert(SaveFilePath.IndexOf("."), " CallSummary")
                SaveMsg = Xl_file
                IsChkBothFiles = True
            End If
            newXlWorkbook.Close()
            newXlWorkbooks = Nothing
            newXlApp.Quit()
            newXlApp = Nothing
            releaseObject(newXlWorkSheet)
            releaseObject(newXlWorkbook)
            releaseObject(newXlWorkbooks)
            releaseObject(newXlApp)
            Exit Sub
        End Try
    End Sub
    Public newXlApp As Excel.Application
    Public newXlWorkbooks As Excel.Workbooks
    Public newXlWorkbook As Excel.Workbook
    Public newXlWorkSheet As Excel.Worksheet
    Public xlSheets As Excel.Worksheets
    Sub CreateNewExcelFile(ByVal FileXLAddress As String)
        newXlApp = New Excel.Application
        newXlWorkbook = newXlApp.Workbooks.Open(FileXLAddress)
        'xlWorkbook = xlWorkbooks.Open(Trim(filePathXL))
        'newXlWorkSheet = "bts"
        'Dim worksheets As Excel.Sheets = newXlWorkbook.Worksheets
        'newXlWorkSheet = DirectCast(worksheets.Add(worksheets(2), Type.Missing, Type.Missing, Type.Missing), Excel.Worksheet)
        'newXlWorkSheet = newXlWorkbook.Sheets("Sheet1")
        newXlWorkSheet = CType(newXlWorkbook.Sheets.Add(Count:=1), Excel.Worksheet)
        newXlWorkSheet.Name = "Call Summary"
    End Sub
    Sub formatcdr(ByVal filePathXL As String)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        prgbar_CDR_report.Minimum = 0
        prgbar_CDR_report.Maximum = 5
        prgbar_CDR_report.Value = 0
        prgbar_CDR_report.Visible = True
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Try
            xlWorkbook = xlWorkbooks.Open(Trim(filePathXL))
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "cdr" Then
                    xlSheet = xlWorkbook.Worksheets("cdr")
                    IsSheetRenamed = True
                    Exit For
                ElseIf xlSheet.Name = "bts" Then
                    xlSheet = xlWorkbook.Worksheets("bts")
                    IsSheetRenamed = True
                    Exit For
                End If
            Next xlSheet
            prgbar_CDR_report.Value = 1
            prgbar_CDR_report.Refresh()
            If IsSheetRenamed = False Then
                MsgBox("Please rename the sheet as 'bts' or 'cdr' and browse again", MsgBoxStyle.OkOnly)
                GoTo Incomplete
            End If
            ' xlSheet = xlWorkbook.Worksheets("bts")
            Dim IsFormatChanged As Boolean = False
            Dim ColNumber As Integer = 1
            Dim ColName As String
            Dim ColRange As String
            While IsFormatChanged = False
                If xlSheet.Cells(1, ColNumber).Value = "a" Then
                    ColName = Chr(65 + ColNumber - 1)
                    ColRange = ColName & ":" & ColName
                    xlSheet.Range(ColRange).NumberFormat = "0"
                    IsFormatChanged = True
                End If
                ColNumber = ColNumber + 1
                If ColNumber > 8 Then
                    MsgBox("Please  rename as 'Time' of time column and browse again", MsgBoxStyle.OkOnly)
                    GoTo Incomplete
                End If
            End While
            prgbar_CDR_report.Value = 2
            prgbar_CDR_report.Refresh()
            IsFormatChanged = False
            ColNumber = 1
            While IsFormatChanged = False
                If xlSheet.Cells(1, ColNumber).Value = "b" Then
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
            prgbar_CDR_report.Value = 3
            prgbar_CDR_report.Refresh()
Incomplete:
            xlApp.DisplayAlerts = False
            ' xlSheet.SaveAs(txt_BTS_Path.Text)
            ' xlApp.DisplayAlerts = True
            xlWorkbook.SaveAs(filePathXL, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkbooks = Nothing
            xlWorkbook.Close(True, misValue, misValue)
            xlApp.Quit()
            ' System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) : xlWorkbook = Nothing
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbooks) : xlWorkbooks = Nothing
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet) : xlSheet = Nothing
            releaseObject(xlSheet)
            releaseObject(xlWorkbook)
            releaseObject(xlWorkbooks)
            releaseObject(xlApp)
            prgbar_CDR_report.Value = 4
            prgbar_CDR_report.Refresh()
        Catch ex As Exception
            MsgBox("An Error occured: " & ex.Message)
            xlWorkbooks = Nothing
            xlWorkbook.Close()
            xlApp.Quit()
            releaseObject(xlSheet)
            releaseObject(xlWorkbook)
            releaseObject(xlWorkbooks)
            releaseObject(xlApp)
        End Try
        prgbar_CDR_report.Value = 5
        prgbar_CDR_report.Refresh()
    End Sub

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
    Public SName As String
    Public ActionBtton As String

    Private Sub btn_Search_in_CDRs_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_in_CDRs.Click
        'prgbar_CDR_report.ForeColor = Color.LightYellow
        'Dim PGB_X As Single = (prgbar_CDR_report.Width / 2) - (frm_Spy_Tech.Width / 2)
        'Dim PGB_Y As Single = (prgbar_CDR_report.Height / 2) - (psize.Height / 2)
        'Dim grph As Graphics = prgbar_CDR_report.CreateGraphics
        'grph.DrawString("Continue work",prgbar_CDR_report.DefaultFont,Brushes.Black,
        'Call ColumnsName()

        Dim XLfileAddress As String = Nothing
        If txt_CDR_Report_Path.Text = "" Then
            Exit Sub
        Else
            Me.Cursor = Cursors.WaitCursor
            frmCdrBTS.ShowDialog()
            If ActionBtton = "Cancel" Then
                Me.Cursor = Cursors.Default
                Exit Sub
            ElseIf ActionBtton = "Yes" Then
                Dim fileExtension As String = Path.GetExtension(Trim(txt_CDR_Report_Path.Text))

                If fileExtension = ".csv" Then
                    'csvFormat(Trim(txt_CDR_Report_Path.Text))
                    txt_CDR_Report_Path.Text = csvTOxlsx(Trim(txt_CDR_Report_Path.Text))
                End If

                'If IsFileCorrect(Trim(txt_CDR_Report_Path.Text), SName) = False Then
                '    Me.Cursor = Cursors.Default
                '    Exit Sub
                'End If
                subsInfo = IsFileCorrect(Trim(txt_CDR_Report_Path.Text), SName)
                If subsInfo(0) <> "" And subsInfo(1) <> "" Then
                    subsInfo(0) = "92" & subsInfo(0).Substring(subsInfo(0).Length - 10, 10)
                    Dim cnic As String = subsInfo(1)
                End If
                Me.Cursor = Cursors.WaitCursor
                Call ChangeColsName(Trim(txt_CDR_Report_Path.Text), SName)

                End If

                Call formatcdr(Trim(txt_CDR_Report_Path.Text))
                XLfileAddress = InsertColsCNIC(Trim(txt_CDR_Report_Path.Text))
                Me.Cursor = Cursors.Default

        End If

        'lbProgress.Text = "Intializing...."

        IsChkBothFiles = False
        If chkOnlyCallSummary.Checked = True Then
            Me.Cursor = Cursors.WaitCursor
            Call CallSammary(XLfileAddress)
            Me.Cursor = Cursors.Default
            Exit Sub
        ElseIf chkBothFiles.Checked = True Then
            Me.Cursor = Cursors.WaitCursor
            Call CallSammary(XLfileAddress)
        End If
        IsCNIC_Need = False
        Me.Cursor = Cursors.WaitCursor
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        xlApp = New Excel.Application
        xlWorkbooks = xlApp.Workbooks
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim IsSheetRenamed As Boolean = False
        Dim SheetName As String = Nothing
        Try
            xlWorkbook = xlWorkbooks.Open(txt_CDR_Report_Path.Text)
            For Each xlSheet In xlWorkbook.Worksheets
                If xlSheet.Name = "bts" Then
                    SheetName = "bts$"
                    IsSheetRenamed = True
                    Exit For
                ElseIf xlSheet.Name = "cdr" Then
                    SheetName = "cdr$"
                    IsSheetRenamed = True
                    Exit For
                End If

            Next xlSheet

        Catch ex As Exception
            MsgBox("Please check the sheet name; sheet name must be 'cdr' or 'bts'", MsgBoxStyle.Information)
            xlWorkbook.Close()
            xlWorkbooks.Close()
            xlApp.Quit()
            Exit Sub
        End Try

        xlWorkbook.Close()
        xlWorkbooks.Close()
        xlApp.Quit()
        If IsSheetRenamed = False Then
            MsgBox("Please check the sheet name; sheet name must be 'cdr' or 'bts'", MsgBoxStyle.Information)

            Exit Sub
        End If


        DT_PhoneNumber.Clear()
        DVG_SpyTech.Rows.Clear()
        Lstbx_CNIC_PhoneNumbers.Items.Clear()
        Call CreateDocTitle("Analysis of: " & DocTitle)
        Dim NextLine As Integer = 1
        DVG_SpyTech.Visible = False
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim ColumnName As String = "a"
        'Dim SheetName As String = "cdr$"
        Dim totalRows As Integer = 0
        Dim queryString As String = "select [" & ColumnName & "] as party, count(*) as CountOf from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        ' Dim queryString As String = "select [" & ColumnName & "] as party, count(*) as CountOf from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        Dim CountqueryString As String = "select COUNT(*) as CountOf from (select [" & ColumnName1 & "] as Party from [" & SheetName & "] GROUP BY [" & ColumnName1 & "]" + _
                        "UNION ALL select [" & ColumnName2 & "] as Party from [" & SheetName & "] GROUP BY [" & ColumnName2 & "] )"
        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & txt_CDR_Report_Path.Text & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
        Dim o As New OleDb.OleDbConnection(ConnectionString)
        o.Open()
        Dim ocmd As New OleDb.OleDbCommand(queryString, o)
        Dim ocmd1 As New OleDb.OleDbCommand(CountqueryString, o)
        Dim oreader As OleDb.OleDbDataReader
        Dim oreader1 As OleDb.OleDbDataReader
        oreader1 = ocmd1.ExecuteReader()
        With oreader1.Read
            totalRows = oreader1(0) + 20
        End With
        'prgbar_CDR_report.BackColor = Color.Transparent
        prgbar_CDR_report.Minimum = 0
        prgbar_CDR_report.Maximum = totalRows
        prgbar_CDR_report.Value = 0
        prgbar_CDR_report.Visible = True
        Dim PhoneNumber As String = Nothing
        Dim IsNumber As Long = Nothing
        ' Dim NumberOfCalls As String = Nothing
        oreader = ocmd.ExecuteReader()
        Dim NumberofRows As String = Nothing
        'Dim Isfound As Boolean = False
        Dim IsFirstNotFound As Boolean = True
        While oreader.Read
            DVG_SpyTech.Rows.Clear()
            prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1
            If oreader.IsDBNull(0) Then
            Else
                PhoneNumber = Trim(oreader("party"))
                Try
                    IsNumber = Long.Parse(PhoneNumber)
                Catch ex As Exception
                    GoTo nextNumber
                End Try



                If (PhoneNumber.Length >= 9) And (PhoneNumber.Length <= 12) Then
                    IsCDR = True
                    NumbersOfCalls = oreader(1).ToString & "," & ColumnName
                    'If getTableName(PhoneNumber) = "Masterdb2021" Then
                    '    Call FindNumberDB2021(PhoneNumber)
                    '    'Call FindNumberDB2022(txt_Phone_Number.Text.Trim)
                    'ElseIf getTableName(PhoneNumber) = "Masterdb2022" Then
                    '    Call FindNumberDB2221(PhoneNumber)
                    'Else
                    '    Call FindNumberDB2022(PhoneNumber)
                    'End If
                    Call FindNumberDB2021(PhoneNumber)
                    ''Call FindPhoneNumber(PhoneNumber)

                    'Following function is Disabled for Rasheed Police
                    Call FindFromOthers(, PhoneNumber)

                    If IsNumberFoundCDR = False Then
                        'If IsFirstNotFound = True Then

                        '    NumbersNotFound = PhoneNumber & "(" & NumbersOfCalls & ")"
                        '    IsFirstNotFound = False
                        'Else
                        '    NumbersNotFound = NumbersNotFound & vbTab & vbTab & vbTab & vbTab & PhoneNumber & "(" & NumbersOfCalls & ")" & vbTab & vbTab
                        'End If
                        If NextLine <= 2 Then
                            If NextLine = 1 Then
                                NumbersNotFound = NumbersNotFound & PhoneNumber & "(" & NumbersOfCalls & ")" & vbTab & vbTab & vbTab & vbTab & vbTab
                                NextLine = 2
                            ElseIf NextLine = 2 Then
                                NumbersNotFound = NumbersNotFound & PhoneNumber & "(" & NumbersOfCalls & ")" & vbCrLf
                                NextLine = 1
                            End If
                        End If
                    End If
                    IsNumberFoundCDR = False
                End If
            End If
nextNumber:
        End While
        NumbersNotFound = NumbersNotFound & vbCrLf
        oreader.Close()
        ColumnName = "b"
        queryString = "select [" & ColumnName & "], count(*) as CountOf from [" & SheetName & "] GROUP BY [" & ColumnName & "]"
        Dim bCommand As New OleDb.OleDbCommand(queryString, o)
        oreader = bCommand.ExecuteReader()
        IsFirstNotFound = True
        While oreader.Read
            prgbar_CDR_report.Value = prgbar_CDR_report.Value + 1
            If oreader.IsDBNull(0) Then
            Else
                PhoneNumber = Trim(oreader(0))
                Try
                    IsNumber = Long.Parse(PhoneNumber)
                Catch ex As Exception
                    GoTo nextNumber1
                End Try
                If (PhoneNumber.Length >= 9) And (PhoneNumber.Length <= 12) Then
                    IsCDR = True
                    NumbersOfCalls = oreader(1).ToString & "," & ColumnName

                    Call FindNumberDB2021(PhoneNumber)
                    ''Call FindPhoneNumber(PhoneNumber)
                    'Following function is disabled for Rasheed Police
                    Call FindFromOthers(, PhoneNumber)
                    If IsNumberFoundCDR = False Then
                        'If IsFirstNotFound = True Then

                        '    NumbersNotFound = NumbersNotFound & PhoneNumber & "(" & NumbersOfCalls & ")"
                        '    IsFirstNotFound = False
                        'Else
                        '    NumbersNotFound = NumbersNotFound & "    ,    " & PhoneNumber & "(" & NumbersOfCalls & ")"
                        'End If
                        If NextLine <= 2 Then
                            If NextLine = 1 Then
                                NumbersNotFound = NumbersNotFound & PhoneNumber & "(" & NumbersOfCalls & ")" & vbTab & vbTab & vbTab & vbTab & vbTab
                                NextLine = 2
                            ElseIf NextLine = 2 Then
                                NumbersNotFound = NumbersNotFound & PhoneNumber & "(" & NumbersOfCalls & ")" & vbCrLf
                                NextLine = 1
                            End If
                        End If

                    End If
                    IsNumberFoundCDR = False
                End If
            End If
NextNumber1:
        End While
        'DT_PhoneNumber.Rows.Add("", "Numbers Not found", NumbersNotFound, Nothing)
        NumbersNotFound = "                                                  Numbers Not Found" & vbCrLf & vbCrLf & NumbersNotFound
        Call PopulateDVG(NumbersNotFound, "", Nothing)
        Call AddRecordInTable(, NumbersNotFound)
        IsCDR = False
        NumbersOfCalls = Nothing
        'DVG_SpyTech.Visible = True
        ' txt_CNC.Text = DVG_SpyTech.Columns(1).Width
        prgbar_CDR_report.Visible = False
        prgbar_CDR_report.Value = 0
        ' PathForDoc = System.IO.Path.GetDirectoryName(PathForDoc)
        PathForDoc = System.IO.Path.GetDirectoryName(Trim(txt_CDR_Report_Path.Text))
        Call DirectSaveDoc(PathForDoc, DocTitle)
        ' Call exportAsWordDoc(DocTitle, PathForDoc, True)
        ' Call btn_Save_in_MSWord_Click()
        o.Close()
        Me.Cursor = Cursors.Default
        btn_Save_in_MSWord.Enabled = False
    End Sub

    Private Sub btn_Search_Registeration_Number_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_Registeration_Number.Click
        If txt_Registeration_Number.Text = "" Then
            MsgBox("Please Enter the Registration Number", MsgBoxStyle.OkOnly)
            txt_Registeration_Number.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        If chk_Add_Result.Checked = False Then
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            Call CreateDocTitle("Analysis of: " & txt_Registeration_Number.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Registeration_Number.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
        DVG_SpyTech.Visible = False
        Call FindFromExise(txt_Registeration_Number.Text)

        If DVG_SpyTech.RowCount = 1 Then
            If DVG_SpyTech.Rows(0).Height > DVG_SpyTech.Height Then
                For i As Integer = 1 To 50
                    DVG_SpyTech.Rows.Insert(i)
                Next

            End If
        End If
        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        Me.Cursor = Cursors.Default
        ' DVG_SpyTech.Visible = True
    End Sub

    Private Sub btn_Search_Engine_Number_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_Engine_Number.Click

        If txt_Engine_Number.Text = "" Then
            MsgBox("Please Enter the Engine Number", MsgBoxStyle.OkOnly)
            txt_Engine_Number.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        If chk_Add_Result.Checked = False Then
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            Call CreateDocTitle("Analysis of: " & txt_Engine_Number.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Engine_Number.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
       
        DVG_SpyTech.Visible = False
        Call FindFromExise(, , txt_Engine_Number.Text)
        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        Me.Cursor = Cursors.Default
        'DVG_SpyTech.Visible = True
    End Sub

    Private Sub btn_Search_Chasis_Number_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_Chasis_Number.Click
        If txt_Chasis_Number.Text = "" Then
            MsgBox("Please Enter the Chasis Number", MsgBoxStyle.OkOnly)
            txt_Chasis_Number.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        If chk_Add_Result.Checked = False Then
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            Call CreateDocTitle("Analysis of: " & txt_Chasis_Number.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Chasis_Number.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
        DVG_SpyTech.Visible = False
        Call FindFromExise(, txt_Chasis_Number.Text)
        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        Me.Cursor = Cursors.Default
        'DVG_SpyTech.Visible = True
    End Sub
    Public AllCommonFileNames() As String
    Public OnlyFilesPath As String
    Private Sub btn_CDRs_CommonLinks_Path_Click(sender As System.Object, e As System.EventArgs) Handles btn_CDRs_CommonLinks_Path.Click
        Dim OpenFileDialog1 As New OpenFileDialog

        Dim FileExtention As String = Nothing
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Excel Files|*.xlsx;*.xls"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim AllFileNames() As String = OpenFileDialog1.FileNames()
            Dim NumberOfFiles As Integer = AllFileNames.Length()
            txt_CDR_CommonLinks_Path.Text = NumberOfFiles & " files have been selected"
            If NumberOfFiles < 2 Then
                MsgBox("Please select more than one files to compare", vbOKOnly)
                btn_Search_Common_Links.Enabled = False
                Exit Sub
            Else
                btn_Search_Common_Links.Enabled = True
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
            'txt_CDR_CommonLinks_Path.Text = PathAndFileName
            'frmXLFormat.NoOfFiles = AllCommonFileNames
            'frmXLFormat.ShowDialog()
        Else
            Exit Sub
        End If
    End Sub

    Sub DropTbl(ByVal tblName As String)
        'Try
        '    Call ConnectionOpenDelDupli()
        '    QueryString = "IF OBJECT_ID('dbo." & tblName & "') IS NOT NULL DROP TABLE " & tblName & ""
        '    CreateTbCommand = New SqlCommand(QueryString, TargetConnectionDelDupli)
        '    CreateTbCommand.ExecuteNonQuery()
        '    TargetConnectionDelDupli.Close()
        'Catch ex As Exception
        '    If TargetConnectionDelDupli.State <> ConnectionState.Closed Then
        '        TargetConnectionDelDupli.Close()
        '    End If
        'End Try
    End Sub

    Private Sub btn_Search_Common_Links_Click(sender As System.Object, e As System.EventArgs) Handles btn_Search_Common_Links.Click
        'Creating word Document
        'Call CommonNos(AllCommonFileNames)

       
        If chbCNIC.Checked = True Or chbCNIC.Checked = False Then
            Dim Reply As MsgBoxResult
            Reply = MsgBox("Do you want to create file with 'CNIC'", MsgBoxStyle.YesNoCancel)
            If Reply = MsgBoxResult.Yes Then
                chbCNIC.Checked = True

            ElseIf Reply = MsgBoxResult.Cancel Then
                Exit Sub
            ElseIf Reply = MsgBoxResult.No Then
                chbCNIC.Checked = False
            End If
        End If
        Me.Cursor = Cursors.WaitCursor
        Call Excel_To_SQL(AllCommonFileNames, OnlyFilesPath)
        Me.Cursor = Cursors.Default
        Exit Sub
        DT_PhoneNumber.Clear()
        DVG_SpyTech.Rows.Clear()
        Lstbx_CNIC_PhoneNumbers.Items.Clear()
        prgbar_Common_Links.Minimum = 0
        prgbar_Common_Links.Maximum = AllCommonFileNames.Length * 3 + 2
        prgbar_Common_Links.Value = 0
        prgbar_Common_Links.Visible = True
        Dim WordApp As New Word.Application()
        Dim doc As New Word.Document()
        doc = WordApp.Documents.Add()
        Dim CommonNoTable As Word.Table
        With doc.Range
            .InsertAfter("Common Numbers Report")
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
        If TotalNumberOfCDRs > 5 Then
            doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
        End If
        prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
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
        CommonNoTable.Rows(1).Cells(2).Range.Text = "Common" & vbCrLf & "Numbers"
        CommonNoTable.Rows(1).Range.Font.Bold = True
        'CommonNoTable.Cell(1, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20
        'CommonNoTable.Cell(1, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20
        'CommonNoTable.Rows.Item(1).Range.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        'CommonNoTable.Rows.Item(1).Range.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic
        'CommonNoTable.Rows.Item(1).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray30
        'CommonNoTable.Columns(2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10
        'CommonNoTable.Columns.AutoFit()
        'CommonNoTable.Range.ParagraphFormat.LeftIndent=
        prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
        Dim ColumnName1 As String = "a"
        Dim ColumnName2 As String = "b"
        Dim SheetName As String = "cdr$"
        Try


            'Create DataTable for results
            OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            OthersConnection.Open()
            Dim TargetConnection As SqlConnection
            TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;")
            TargetConnection.Open()
            Dim Delquery As String = "DELETE FROM CommonNumbers"
            OthersCommand = New SqlCommand(Delquery, OthersConnection)
            OthersCommand.ExecuteNonQuery()
            Dim CreateTablesQuery As String
            Dim TempTableName As String
            Dim QueryInsert As String
            Dim queryString1 As String
            For j As Integer = 0 To TotalNumberOfCDRs - 1
                prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
                OnlyFileNames(j) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(j))
                TempTableName = "_" & OnlyFileNames(j)
                CommonNoTable.Rows(1).Cells(j + 3).Range.Text = OnlyFileNames(j)
                'CommonNoTable.Cell(1, j + 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20
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
                'CreateTablesQuery = "CREATE CLUSTERED INDEX myIdx ON  [" & TempTableName & "](PhoneNumber)"
                'OthersCommand = New SqlCommand(CreateTablesQuery, OthersConnection)
                'Try
                '    OthersCommand.ExecuteNonQuery()
                'Catch ex02 As Exception
                '    MsgBox("creating table", MsgBoxStyle.OkOnly)
                'End Try
                Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=0;IMEX=0'"
                Dim o As New OleDb.OleDbConnection(ConnectionString)
                o.Open()
                'populate table with data with 92
                Dim queryString As String = "select party from (select [" & ColumnName1 & "] as Party from [" & SheetName & "]" + _
                            "UNION ALL select [" & ColumnName2 & "] as Party from [" & SheetName & "])"
                Dim InsertCommand As New OleDb.OleDbCommand(queryString, o)
                Dim InsertReader As OleDb.OleDbDataReader
                Try
                    InsertReader = InsertCommand.ExecuteReader()
                Catch ex02 As Exception
                    MsgBox("insertion error table", MsgBoxStyle.OkOnly)
                End Try
                While InsertReader.Read
                    If InsertReader.IsDBNull(0) Then
                    Else
                        Dim ConsistantNumber As String = InsertReader(0)
                        ConsistantNumber = Trim(ConsistantNumber)
                        'ConsistantNumber = StandardNumber(ConsistantNumber)
                        If ConsistantNumber.Length >= 9 Then
                            If ConsistantNumber.Substring(0, 2) <> "92" Then
                                If ConsistantNumber.Substring(0, 1) = "0" Then
                                    If ConsistantNumber.Length = 11 Then
                                        ConsistantNumber = "92" & ConsistantNumber.Substring(1, 10)
                                    ElseIf ConsistantNumber.Length = 10 Then
                                        ConsistantNumber = "92" & ConsistantNumber.Substring(1, 9)
                                    End If
                                ElseIf ConsistantNumber.Substring(0, 1) <> "0" And ConsistantNumber.Length >= 9 Then
                                    ConsistantNumber = "92" & ConsistantNumber
                                End If
                            End If
                            'ConsistantNumber = StandardNumber(ConsistantNumber)
                            Try
                                QueryInsert = "INSERT INTO CommonNumbers  Values ('" & ConsistantNumber & "')"
                                OthersCommand = New SqlCommand(QueryInsert, OthersConnection)
                                OthersCommand.ExecuteNonQuery()
                                QueryInsert = "INSERT INTO [" & TempTableName & "]  Values ('" & ConsistantNumber & "')"
                                OthersCommand = New SqlCommand(QueryInsert, OthersConnection)
                                OthersCommand.ExecuteNonQuery()
                            Catch ex01 As Exception
                                MsgBox("Error during insertion data & vbCrLf", vbOKOnly)
                            End Try


                        End If
                    End If
                End While
                InsertReader.Close()
                o.Close()
            Next
            'Test Area
            ' '' '' ''For j As Integer = 0 To TotalNumberOfCDRs - 1
            ' '' '' ''    Dim ConnectionString1 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
            ' '' '' ''    Dim o1 As New OleDb.OleDbConnection(ConnectionString1)
            ' '' '' ''    o1.Open()
            ' '' '' ''    Dim queryStringUpdate As String = "select a from [" & SheetName & "] where NOT LIKE '" & "92%" & "'"
            ' '' '' ''    Dim InsertCommand1 As New OleDb.OleDbCommand(queryStringUpdate, o1)
            ' '' '' ''    Dim InsertReader1 As OleDb.OleDbDataReader
            ' '' '' ''    InsertReader1 = InsertCommand1.ExecuteReader()
            ' '' '' ''    Dim UpdateQuery As String
            ' '' '' ''    Dim NewNumber As Integer
            ' '' '' ''    Dim oldNumber As String
            ' '' '' ''    While InsertReader1.Read
            ' '' '' ''        If InsertReader1.IsDBNull(0) Then
            ' '' '' ''        Else
            ' '' '' ''            oldNumber = InsertReader1(0)
            ' '' '' ''            If oldNumber.Substring(0, 1) = "0" Then
            ' '' '' ''                If oldNumber.Length = 11 Then
            ' '' '' ''                    NewNumber = CType("92" & oldNumber.Substring(1, 10), Integer)
            ' '' '' ''                ElseIf oldNumber.Length = 10 Then
            ' '' '' ''                    NewNumber = CType("92" & oldNumber.Substring(1, 9), Integer)
            ' '' '' ''                End If
            ' '' '' ''            Else
            ' '' '' ''                NewNumber = CType("92" & oldNumber, Integer)
            ' '' '' ''                 UpdateQuery = "UDATE [" & SheetName & "] SET a = " & NewNumber & " Where a= " & oldNumber & ""
            ' '' '' ''            End If
            ' '' '' ''            Dim ConnectionString2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & AllCommonFileNames(j) & ";Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=0'"
            ' '' '' ''            Dim o2 As New OleDb.OleDbConnection(ConnectionString1)
            ' '' '' ''            UpdateQuery = "UDATE [" & SheetName & "] SET a = " & NewNumber & " Where a= " & oldNumber & ""
            ' '' '' ''            Dim InsertCommand2 As New OleDb.OleDbCommand(UpdateQuery, o2)
            ' '' '' ''            InsertCommand2.ExecuteNonQuery()
            ' '' '' ''        End If
            ' '' '' ''    End While
            ' '' '' ''Next
            'Test Area 
            'Call CreateDTCommonNumbers(TotalNumberOfCDRs, OnlyFileNames)
            Dim isCounterSet As Boolean = False
            Dim conmonNoCount As String = "Select COUNT ( DISTINCT PhoneNumber ) AS Counter FROM CommonNumbers"
            OthersCommand = New SqlCommand(conmonNoCount, OthersConnection)
            OthersReader = OthersCommand.ExecuteReader()
            prgbar_Common_Links.Value = 0
            While OthersReader.Read
                If isCounterSet = False Then
                    prgbar_Common_Links.Maximum = OthersReader("Counter") + 10
                    isCounterSet = True
                End If
            End While
            OthersReader.Close()
            'OthersConnection.Close()
            'Dim CommonNoQuery As String = "SELECT PhoneNumber FROM CommonNumbers GROUP BY PhoneNumber"
            Dim CommonNoQuery As String = "SELECT DISTINCT PhoneNumber FROM CommonNumbers"
            OthersCommand = New SqlCommand(CommonNoQuery, OthersConnection)
            OthersReader = OthersCommand.ExecuteReader()

            Dim PhoneNumber As String = Nothing
            Dim NumberofRows As String = Nothing
            Dim IsFound As Boolean = False
            Dim RecCounter As Integer
            Dim RowIndex As Integer = 2
            prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
            Dim tableNames As New List(Of String)()
            Dim Select_string As String = Nothing
            Dim From_string As String = "From"
            Dim PhoneNum As String = "MobileNum"

            'For i As Integer = 0 To TotalNumberOfCDRs - 1

            '    Dim PathAndFileName As String = AllCommonFileNames(i)
            '    OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
            '    Dim currentFileName = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
            '    ' tableNames.Add("_" & OnlyFileNames(i))
            '    TempTableName = "_" & OnlyFileNames(i)
            '    PhoneNumber = "923004994213"
            '    Select_string = Select_string + "(select  count(*) from  [" & TempTableName & "] as [" & TempTableName & "] WHERE PhoneNumber = '" & PhoneNum & "'),"
            '    'From_string = From_string + " _" & OnlyFileNames(i) + ","
            '    'Select_string = Select_string + " Count(distinct"
            'Next

            'Select_string = "Select" + Select_string.Substring(0, Select_string.Length - 1)
            ''From_string = From_string.Substring(0, From_string.Length - 1)

            'Dim NewTargetQuery As String = Select_string.Replace("MobileNum", PhoneNumber)
            'Dim TargetFileCommand1 As SqlCommand
            'TargetFileCommand1 = New SqlCommand(NewTargetQuery, TargetConnection)
            'Dim TargetReader1 As SqlDataReader
            'TargetReader1 = TargetFileCommand1.ExecuteReader
            'While TargetReader1.Read()
            '    Dim count1 As String = TargetReader1(0).ToString
            '    Dim count2 As String = TargetReader1(1).ToString
            '    Dim count3 As String = TargetReader1(2).ToString
            'End While
            Dim currentFileName As String
            Dim PathAndFileName As String
            Dim TargetFileCommand As SqlCommand
            Dim TargetReader As SqlDataReader
            Dim SelectNewString As String
            For i As Integer = 0 To TotalNumberOfCDRs - 1

                PathAndFileName = AllCommonFileNames(i)
                OnlyFileNames(i) = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                currentFileName = System.IO.Path.GetFileNameWithoutExtension(AllCommonFileNames(i))
                TempTableName = "_" & OnlyFileNames(i)
                Select_string = Select_string + "(select  count(*) from  [" & TempTableName & "] as [" & TempTableName & "] WHERE PhoneNumber = '" & PhoneNum & "'),"

                'Dim TargetFileQueryString As String = "select PhoneNumber, count(*) as CountOf from  [" & TempTableName & "]  WHERE PhoneNumber = '" & PhoneNumber & "'  GROUP BY PhoneNumber"
                'OthersConnection.Open()
                'Dim TargetFileQueryString As String = "select PhoneNumber, count(*) as CountOf from  " & TempTableName & "  WHERE PhoneNumber LIKE '" & "%" & PhoneNumber & "%" & "'  GROUP BY PhoneNumber"
            Next
            Select_string = "Select " + Select_string.Substring(0, Select_string.Length - 1)
            prgbar_Common_Links.Value = 0
            While OthersReader.Read
                ' NumberOfCommons = Nothing
                RecCounter = 0

                '   DT_CommonNumbers.Rows.Add()
                CommonNoTable.Rows.Add()
                'DT_CommonNumbers.Rows(RowIndex).Item(0) = OthersReader(0).ToString
                If OthersReader.IsDBNull(0) Then

                Else
                    'Finding CNIC of phone number

                    PhoneNumber = OthersReader(0).ToString
                    If isCounterSet = False Then
                        ' prgbar_Common_Links.Maximum = OthersReader("Counter") + 10
                        isCounterSet = True
                    End If



                    'IsCNIC_Need = True
                    'GprCNIC = Nothing
                    'If chbCNIC.Checked = True Then
                    '    Call FindNumberDB2021(PhoneNumber)
                    '    '' Call FindPhoneNumber(PhoneNumber)
                    'End If
                    'If GprCNIC = Nothing Then
                    '    GprCNIC = "NA"
                    'End If
                    'IsCNIC_Need = False




                    SelectNewString = Select_string.Replace("MobileNum", PhoneNumber)
                    TargetFileCommand = New SqlCommand(SelectNewString, TargetConnection)

                    TargetReader = TargetFileCommand.ExecuteReader
                    While TargetReader.Read()

                        ' NumberOfCommons(i) = TargetReader(i).ToString
                        '           DT_CommonNumbers.Rows(RowIndex).Item(OnlyFileNames(i)) = NumberOfCommons(i)
                        With CommonNoTable
                            .Rows(RowIndex).Cells(1).Range.Text = RowIndex - 1

                            For i As Integer = 0 To TotalNumberOfCDRs - 1
                                If TargetReader(i).ToString <> "0" Then
                                    .Rows(RowIndex).Cells(i + 3).Range.Text = TargetReader(i)
                                    RecCounter = RecCounter + 1
                                End If
                            Next
                            If RecCounter > 1 Then
                                IsCNIC_Need = True
                                GprCNIC = Nothing
                                If chbCNIC.Checked = True Then
                                    Call FindNumberDB2021(PhoneNumber)
                                    '' Call FindPhoneNumber(PhoneNumber)
                                End If
                                IsCNIC_Need = False
                                If GprCNIC = Nothing Then
                                    .Rows(RowIndex).Cells(2).Range.Text = PhoneNumber
                                Else
                                    .Rows(RowIndex).Cells(2).Range.Text = PhoneNumber & vbCrLf & GprCNIC
                                End If
                            End If
                        End With


                        '  MsgBox("Number has been found in " + TargetReader(0).ToString + " " + TargetReader(1).ToString, vbOKOnly)
                    End While


                    TargetReader.Close()



                    If RecCounter <= 1 Then
                        '      DT_CommonNumbers.Rows.RemoveAt(RowIndex)
                        CommonNoTable.Rows(RowIndex).Delete()
                    Else

                        RowIndex = RowIndex + 1
                    End If
                End If
                prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
            End While
            'NumbersNotFound = NumbersNotFound & vbCrLf & vbCrLf
            OthersReader.Close()
            For i As Integer = 0 To TotalNumberOfCDRs - 1
                prgbar_Common_Links.Value = prgbar_Common_Links.Value + 1
                TempTableName = "_" & OnlyFileNames(i)
                Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
                OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
                OthersCommand.ExecuteNonQuery()
            Next
            CommonNoTable.Columns.AutoFit()
        Catch ex As Exception
            For i As Integer = 0 To TotalNumberOfCDRs - 1
                Try
                    Dim TempTableName As String = "_" & OnlyFileNames(i)
                    Dim TableDropQuery As String = "Drop Table [" & TempTableName & "]"
                    OthersCommand = New SqlCommand(TableDropQuery, OthersConnection)
                    OthersCommand.ExecuteNonQuery()
                Catch ex1 As Exception
                    GoTo nextItration
                End Try
nextItration:
            Next

            MsgBox("The result has not been produced" & vbCrLf & "Please make sure sheet is 'cdr' with columns name a and b", vbOKOnly)
            prgbar_Common_Links.Visible = False
            prgbar_Common_Links.Value = 0
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try

        Try
            WordApp.Options.SavePropertiesPrompt = False
            doc.SaveAs(OnlyFilesPath & "\CommonNumberReport.docx")
            doc.Close(Word.WdSaveOptions.wdSaveChanges)

            prgbar_Common_Links.Value = AllCommonFileNames.Length * 3 + 2
            prgbar_Common_Links.Visible = False
            prgbar_Common_Links.Value = 0
            MsgBox("The File has been created: " & vbCrLf & OnlyFilesPath & "\CommonNumbersReport", vbOKOnly)
        Catch ex As Exception
            'Dim MSWord As New Word.Application
            'Dim WordDoc As New Word.Document
            prgbar_Common_Links.Visible = False
            prgbar_Common_Links.Value = 0
            If File.Exists(OnlyFilesPath & "\CommonNumberReport.docx") Then
                doc.Close(Word.WdSaveOptions.wdSaveChanges)

            End If
            'doc.SaveAs(OnlyFilesPath & "\CommonNumberReport.docx")
            'doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            'MsgBox("The File has been created: " & vbCrLf & OnlyFilesPath & "\CommonNumbersReport", vbOKOnly)
        End Try
        Me.Cursor = Cursors.Default
        btn_Save_in_MSWord.Enabled = False
        'OthersConnection.Close()




    End Sub

    Private Sub txt_CDR_Report_Path_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_CDR_Report_Path.KeyPress
        Beep()
        e.Handled = True
    End Sub

    Private Sub txt_CDR_Report_Path_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_CDR_Report_Path.TextChanged

    End Sub

    Private Sub txt_CDR_CommonLinks_Path_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_CDR_CommonLinks_Path.KeyPress
        Beep()
        e.Handled = True
    End Sub

    Private Sub txt_CDR_CommonLinks_Path_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_CDR_CommonLinks_Path.TextChanged

    End Sub

    Private Sub btn_Employee_Name_Click(sender As System.Object, e As System.EventArgs) Handles btn_Employee_Name.Click
        If txt_Employee_Name.Text = "" Then
            MsgBox("Please Enter the Name", MsgBoxStyle.OkOnly)
            txt_Employee_Name.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        If chk_Add_Result.Checked = False Then
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            Call CreateDocTitle("Analysis of: " & txt_Employee_Name.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Employee_Name.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
        DVG_SpyTech.Visible = False
        Call FindFromEmployee(, txt_Employee_Name.Text)

        

        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btn_Employee_CNIC_Click(sender As System.Object, e As System.EventArgs) Handles btn_Employee_CNIC.Click
        If txt_Employee_CNIC.Text = "" Or txt_Employee_CNIC.Text.Length < 13 Then
            MsgBox("Please Enter the complete 13 digites of CNIC", MsgBoxStyle.OkOnly)
            txt_Employee_CNIC.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        If chk_Add_Result.Checked = False Then
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            Call CreateDocTitle("Analysis of: " & txt_Employee_CNIC.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Employee_CNIC.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
        DVG_SpyTech.Visible = False
        Call FindFromEmployee(, , txt_Employee_CNIC.Text)



        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btn_Employee_Belt_No_Click(sender As System.Object, e As System.EventArgs) Handles btn_Employee_Belt_No.Click
        If txt_Employee_Belt_No.Text = "" Then
            MsgBox("Please Enter the Belt Number", MsgBoxStyle.OkOnly)
            txt_Employee_Belt_No.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        If chk_Add_Result.Checked = False Then
            DT_PhoneNumber.Clear()
            DVG_SpyTech.Rows.Clear()
            Lstbx_CNIC_PhoneNumbers.Items.Clear()
            Call CreateDocTitle("Analysis of: " & txt_Employee_Belt_No.Text)
        Else
            If isDocSaved = True Then
                Call CreateDocTitle("Analysis of: " & txt_Employee_Belt_No.Text)
                DT_PhoneNumber.Clear()
                DVG_SpyTech.Rows.Clear()
                isDocSaved = False
            End If
        End If
        DVG_SpyTech.Visible = False
        Call FindFromEmployee(txt_Employee_Belt_No.Text)



        If DVG_SpyTech.RowCount > 0 Then
            DVG_SpyTech.Visible = True
            btn_Save_in_MSWord.Enabled = True
        Else
            MsgBox("Not found", MsgBoxStyle.OkOnly)
            DVG_SpyTech.Visible = False
            btn_Save_in_MSWord.Enabled = False
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub txt_Employee_Name_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Employee_Name.KeyPress
        If e.KeyChar = vbCr Then
            btn_Employee_Name_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Employee_Name_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Employee_Name.TextChanged

    End Sub

    Private Sub txt_CNC_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_CNC.KeyPress
        If e.KeyChar = vbCr Then
            btn_Search_CNC_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_CNC_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_CNC.TextChanged

    End Sub

    Private Sub txt_Registeration_Number_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Registeration_Number.KeyPress
        If e.KeyChar = vbCr Then
            btn_Search_Registeration_Number_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Registeration_Number_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Registeration_Number.TextChanged

    End Sub

    Private Sub txt_Engine_Number_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Engine_Number.KeyPress
        If e.KeyChar = vbCr Then
            btn_Search_Engine_Number_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Engine_Number_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Engine_Number.TextChanged

    End Sub

    Private Sub txt_Chasis_Number_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Chasis_Number.KeyPress
        If e.KeyChar = vbCr Then
            btn_Search_Chasis_Number_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Chasis_Number_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Chasis_Number.TextChanged

    End Sub

    Private Sub txt_Employee_CNIC_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Employee_CNIC.KeyPress
        If e.KeyChar = vbCr Then
            btn_Employee_CNIC_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Employee_CNIC_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Employee_CNIC.TextChanged

    End Sub

    Private Sub txt_Employee_Belt_No_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Employee_Belt_No.KeyPress
        If e.KeyChar = vbCr Then
            btn_Employee_Belt_No_Click(sender, e)
            e.Handled = True
        End If
    End Sub

    Private Sub txt_Employee_Belt_No_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Employee_Belt_No.TextChanged

    End Sub

    Private Sub btnBTS_Click(sender As System.Object, e As System.EventArgs) Handles btnBTS.Click
        Form1.Show()
    End Sub

    
    Private Sub chkOnlyCallSummary_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkOnlyCallSummary.CheckedChanged
        If chkOnlyCallSummary.Checked = True Then
            chkBothFiles.Checked = False
        End If
    End Sub

    Private Sub chkBothFiles_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBothFiles.CheckedChanged
        If chkBothFiles.Checked = True Then
            chkOnlyCallSummary.Checked = False
        End If
    End Sub


    Private Sub btnMergeExcelFiles_Click(sender As System.Object, e As System.EventArgs) Handles btnMergeExcelFiles.Click
        frmMergeExcels.StartPosition = FormStartPosition.CenterParent
        frmMergeExcels.ShowDialog()

    End Sub

    Private Sub btnMemoryRelease_Click(sender As System.Object, e As System.EventArgs) Handles btnMemoryRelease.Click

        Call UnusedSpace("Others")

        Dim conn As SqlConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") '("Server=" + ServerName + ";Database=MasterDB2022;User Id=sajjad;Password=rajpoot;")
        conn.Open()
        'Dim sql As String = "DBCC SHRINKFILE (Others_data, EMPTYFILE);"
        'Dim cmd As New SqlCommand(sql, conn)
        'Try
        '    Me.Cursor = Cursors.WaitCursor
        '    Dim Result As Integer = cmd.ExecuteScalar()
        '    conn.Close()
        '    'MsgBox("Memory released")
        '    MessageBox.Show("Total " & Result & " kilobytes of space were released.")
        'Catch ex As Exception
        '    conn.Close()
        '    MsgBox(Err.Description)
        'End Try
        'Me.Cursor = Cursors.Default
       

        Try
            Cursor = Cursors.WaitCursor
            Dim command As New SqlCommand("DBCC SHRINKDATABASE (Others, TRUNCATEONLY)", conn)
            Dim Result As Integer = command.ExecuteScalar()
            conn.Close()
            MessageBox.Show("Total " & Result & " kilobytes of space were released.")
        Catch ex As Exception
            conn.Close()
            MsgBox(Err.Description)
        End Try
        Cursor = Cursors.Default

        'Execute the command'


    End Sub
    Sub UnusedSpace(ByVal databaseName As String)
        Dim conn As SqlConnection = New SqlConnection("Server=" + ServerName + ";Database=" + databaseName + ";Trusted_Connection=True;") '("Server=" + ServerName + ";Database=MasterDB2022;User Id=sajjad;Password=rajpoot;")
        conn.Open()
        Dim cmd As New SqlCommand("SELECT * FROM sys.dm_db_file_space_usage", conn)
        Dim reader As SqlDataReader = cmd.ExecuteReader()

        If reader.HasRows Then
            While reader.Read()
                For i As Integer = 0 To reader.FieldCount - 1
                    MessageBox.Show("Column name: " & reader.GetName(i))
                Next
                'Dim file_id As Integer = reader("file_id")
                'Dim total_size As Decimal = reader("total_page_count")
                'Dim used_size As Decimal = reader("used_page_count")
                'Dim unused_size As Decimal = total_size - used_size
                'Console.WriteLine("File ID: " & file_id)
                'Console.WriteLine("Total Size (KB): " & total_size)
                'Console.WriteLine("Used Size (KB): " & used_size)
                'Console.WriteLine("Unused Size (KB): " & unused_size)

            End While
        End If

        reader.Close()
    End Sub

    Private Sub btnVerisys_Click(sender As System.Object, e As System.EventArgs) Handles btnVerisys.Click
        frmVerisys.Show()
    End Sub

   
End Class
