
Public Class LoginForm1
    Public isLoginPass As Boolean = False
   

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        If UsernameTextBox.Text = Nothing Or PasswordTextBox.Text = Nothing Then
            MsgBox("Please enter username and password", MsgBoxStyle.Information)
            UsernameTextBox.Focus()
            UsernameTextBox.SelectAll()
            Exit Sub
        End If
        Dim firstdate As Date = Date.Parse(My.Computer.Registry.CurrentUser.GetValue("IntialDate"))
        Dim lastdate As Date = Date.Parse(My.Computer.Registry.CurrentUser.GetValue("EndDate"))
        Dim days As Long = DateDiff(DateInterval.Day, firstdate, lastdate)
        'Dim expDays As Integer = exampleDays.Days
        Dim password As String = My.Computer.Registry.CurrentUser.GetValue("Password")
        Dim username As String = My.Computer.Registry.CurrentUser.GetValue("UserName")
        'Dim Diff As TimeSpan = Convert.ToDateTime(My.Computer.Registry.CurrentUser.GetValue("EndDate")) - Convert.ToDateTime(My.Computer.Registry.CurrentUser.GetValue("EndDate"))

        'Dim TotalDays As Integer = Diff.Days
        'If days >= 31 Then
        '    Try
        '        My.Computer.Registry.CurrentUser.DeleteSubKey("Password", True)
        '        My.Computer.Registry.CurrentUser.DeleteSubKey("UserName", True)
        '        My.Computer.Registry.CurrentUser.DeleteValue("Password", True)
        '        My.Computer.Registry.CurrentUser.DeleteValue("UserName", True)
        '        My.Computer.Registry.CurrentUser.DeleteSubKey("IntialDate", True)
        '        My.Computer.Registry.CurrentUser.DeleteSubKey("EndDate", True)
        '        My.Computer.Registry.CurrentUser.DeleteValue("IntialDate", True)
        '        My.Computer.Registry.CurrentUser.DeleteValue("EndDate", True)
        '    Catch ex As Exception

        '    End Try
        '    Me.Close()
        '    Exit Sub
        'Else
        My.Computer.Registry.CurrentUser.SetValue("EndDate", Today)
        'End If
        If password = Nothing Or username = Nothing Then
            MsgBox("Software access problem, Please contact the developer", MsgBoxStyle.Information)
            Me.Close()
            'frm_Spy_Tech.Close()

            Exit Sub
        End If
        If UsernameTextBox.Text = username And PasswordTextBox.Text = password Then

            Me.Hide()
            frm_Spy_Tech.Show()
            frm_Spy_Tech.Enabled = True
            frm_Spy_Tech.txt_Phone_Number.Focus()
        Else
            MsgBox("Username or Password is incorrect, Please try again", MsgBoxStyle.Information)
            UsernameTextBox.Focus()
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
        frm_Spy_Tech.Close()
    End Sub

    Private Sub PasswordTextBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles PasswordTextBox.KeyDown

    End Sub

    Private Sub PasswordTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles PasswordTextBox.KeyPress

    End Sub

   

   
    Private Sub PasswordTextBox_TextChanged(sender As System.Object, e As System.EventArgs) Handles PasswordTextBox.TextChanged

    End Sub

    Private Sub UsernameTextBox_TextChanged(sender As System.Object, e As System.EventArgs) Handles UsernameTextBox.TextChanged

    End Sub
End Class
