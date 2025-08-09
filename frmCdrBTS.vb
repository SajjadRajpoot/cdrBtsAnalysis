Public Class frmCdrBTS

    
    Private Sub frmCdrBTS_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.AcceptButton = Button1
        Me.CancelButton = Button2
        Button1.Enabled = False
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        Me.CenterToScreen()
        
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim selectedFile As String
        If RadioButton1.Checked = True Then
            selectedFile = "CDR"
            frm_Spy_Tech.SName = "cdr"
        ElseIf RadioButton2.Checked = True Then
            selectedFile = "BTS"
            frm_Spy_Tech.SName = "bts"
        End If
        Dim response = MsgBox("Are you sure the given file is " & selectedFile, MsgBoxStyle.YesNo)
        If response = MsgBoxResult.Yes Then
            frm_Spy_Tech.ActionBtton = "Yes"
            Me.Close()
        Else
            Exit Sub
        End If


    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Button1.Enabled = True
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Button1.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frm_Spy_Tech.ActionBtton = "Cancel"
    End Sub
End Class