Imports Microsoft.Office.Interop
Module ExportToWord
    Public ExportWordApp As New Word.Application()
    Public ExportIndoc As New Word.Document()
    Public CommonNoTable As Word.Table
    Public InitialIndex As Integer
    Public Initialized As Boolean = False
    Public InitialValue As String = ""
    Public IsDocCreated As Boolean = False

    Function CreateDocTitle(ByVal Title As String)
        If IsDocCreated = True Then
            ExportIndoc=Nothing
            IsDocCreated = False
        End If
        ExportIndoc = ExportWordApp.Documents.Add()
        IsDocCreated = True
        With ExportIndoc.Range
            .InsertAfter(Title)
            .InsertParagraphAfter()
            .InsertAfter("Date:  " & Date.Today & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Time:  " & Date.Now.ToLongTimeString)
            .InsertParagraphAfter()
            .InsertParagraphAfter()
        End With
        With ExportIndoc.PageSetup
            .PaperSize = Word.WdPaperSize.wdPaperA4
            .LeftMargin = 20
            .RightMargin = 20
            .TopMargin = 30
            .BottomMargin = 30
        End With
        Dim SelRange As Word.Range
        SelRange = ExportIndoc.Paragraphs.Item(1).Range
        SelRange.Font.Size = 14
        SelRange.Font.Bold = True
        SelRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        SelRange = ExportIndoc.Paragraphs.Item(2).Range
        SelRange.Font.Size = 12
        SelRange.Font.Bold = True
        SelRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

        'Creating Table in Word Docurment
        CommonNoTable = ExportIndoc.Range.Tables.Add(ExportIndoc.Bookmarks.Item("\endofdoc").Range, 1, 3)
        CommonNoTable.Borders.OutsideColor = Word.WdColor.wdColorBlack
        CommonNoTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        CommonNoTable.Borders.InsideColor = Word.WdColor.wdColorBlack
        CommonNoTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        CommonNoTable.Range.Font.Size = 9
        CommonNoTable.Columns(1).Width = 75
        CommonNoTable.Columns(2).Width = 350
        CommonNoTable.Columns(3).Width = 120
        CommonNoTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
        CommonNoTable.Range.ParagraphFormat.LineSpacing = 1
        CommonNoTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast
        CommonNoTable.Range.ParagraphFormat.SpaceBefore = 0
        CommonNoTable.Range.ParagraphFormat.SpaceAfter = 0
        'CommonNoTable.Style = "Light Grid"
        'CommonNoTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
        'CommonNoTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        CommonNoTable.Rows(1).Cells(1).Range.Text = "Phone" & vbCrLf & "Number"
        CommonNoTable.Rows(1).Cells(1).Range.Font.Bold = True
        CommonNoTable.Rows(1).Cells(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        CommonNoTable.Rows(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
        CommonNoTable.Rows(1).Cells(2).Range.Text = "Detail"
        CommonNoTable.Rows(1).Cells(2).Range.Font.Bold = True
        CommonNoTable.Rows(1).Cells(2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        CommonNoTable.Rows(1).Cells(3).Range.Text = "Photograph"
        CommonNoTable.Rows(1).Cells(3).Range.Font.Bold = True
        CommonNoTable.Rows(1).Cells(3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

    End Function
    Function AddRecordInTable(Optional ByVal PhoneNumber As String = "", Optional ByVal Detail As String = "", Optional ByVal Photograph As Byte() = Nothing)
        CommonNoTable.Rows.Add()
        Dim ImageStream As System.IO.MemoryStream
        Dim _image As Image = Nothing
        If IsNothing(Photograph) = False Then

            ImageStream = New System.IO.MemoryStream(Photograph)
            _image = Image.FromStream(ImageStream)
            'Dim imgwidth As Integer = _image.Width
            'Dim imghight As Integer = _image.Height
            'Dim bmp As New Bitmap(_image)
            'Dim Thumb As New Bitmap(120, 160)
            'Dim g As Graphics = Graphics.FromImage(Thumb)
            'g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            'g.DrawImage(bmp, New Rectangle(0, 0, 120, 160), New Rectangle(0, 0, bmp.Width, bmp.Height), GraphicsUnit.Pixel)
            '_image = DirectCast(Thumb, Image)
            ImageStream.Dispose()
        End If
       

        'imgwidth = _image.Width
        'imghight = _image.Height
        Dim RowIndex As Integer = CommonNoTable.Rows.Count
        ''If Initialized = False Then
        ''    InitialIndex = RowIndex
        ''    Initialized = True
        ''    InitialValue = PhoneNumber
        ''End If
        With CommonNoTable
            .Cell(RowIndex, 1).Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            .Cell(RowIndex, 1).Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            If (InitialValue = PhoneNumber) And (RowIndex > 2) Then
                ' If InitialIndex <> RowIndex - 1 Then
                '.Cell(InitialIndex, 1).Merge(.Cell(RowIndex - 1, 1))
                .Cell(RowIndex, 1).Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                'End If
                'InitialIndex = RowIndex
                'Initialized = True
            Else
                .Rows(RowIndex).Cells(1).Range.Text = PhoneNumber
            End If

            .Rows(RowIndex).Cells(1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
            .Rows(RowIndex).Cells(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

            .Rows(RowIndex).Cells(1).Range.Font.Bold = True
            .Rows(RowIndex).Cells(2).Range.Text = Detail
            .Rows(RowIndex).Cells(2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            .Rows(RowIndex).Cells(1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
            .Rows(RowIndex).Cells(2).Range.Font.Bold = False
            .Rows(RowIndex).Cells(3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Rows(RowIndex).Cells(3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
            If IsNothing(_image) = False Then
                Clipboard.SetImage(_image)
                .Rows(RowIndex).Cells(3).Range.Paste()
                _image.Dispose()
            End If


            InitialValue = PhoneNumber
            Clipboard.Clear()

        End With
    End Function
    Public isDocSaved As Boolean = False
    Public IsChkBothFiles As Boolean = False
    Public SaveMsg As String = Nothing
    Function DirectSaveDoc(ByVal FilePath As String, ByVal FileName As String)
        ExportWordApp.Options.SavePropertiesPrompt = False
        ExportIndoc.SaveAs(FilePath & "\" & FileName & ".docx")
        ExportIndoc.Close(Word.WdSaveOptions.wdSaveChanges)
        frm_Spy_Tech.btn_Save_in_MSWord.Enabled = False
        isDocSaved = True
        IsDocCreated = False
        If IsChkBothFiles = True Then
            MsgBox("File has been created:" & vbCrLf & FilePath & "\" & FileName & ".docx" & vbCrLf & SaveMsg, MsgBoxStyle.Information)
        Else
            MsgBox("File has been created:" & vbCrLf & FilePath & "\" & FileName & ".docx", MsgBoxStyle.OkOnly)
        End If

    End Function
    Function ResizedImage(ByVal Photograph As Byte(), ByVal imgWidth As Integer, ByVal imgHeight As Integer) As Image
        Dim ImageStream As System.IO.MemoryStream
        Dim _image As Image
        ImageStream = New System.IO.MemoryStream(Photograph)
        _image = Image.FromStream(ImageStream)
        'Dim imgwidth As Integer = _image.Width
        'Dim imghight As Integer = _image.Height
        Dim bmp As New Bitmap(_image)
        Dim Thumb As New Bitmap(imgWidth, imgHeight)
        Dim g As Graphics = Graphics.FromImage(Thumb)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(bmp, New Rectangle(0, 0, imgWidth, imgWidth), New Rectangle(0, 0, bmp.Width, bmp.Height), GraphicsUnit.Pixel)
        _image = DirectCast(Thumb, Image)
        'imgwidth = _image.Width
        'imghight = _image.Height
        ImageStream.Dispose()
        'Thumb.Dispose()
        bmp.Dispose()
        Return _image
        Thumb.Dispose()
    End Function
End Module
