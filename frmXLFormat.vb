Public Class frmXLFormat
    Public NoOfFiles() As String
    Private Sub frmXLFormat_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        XLFormat(NoOfFiles)
        'colHeads(NoOfFiles)
    End Sub

    
End Class