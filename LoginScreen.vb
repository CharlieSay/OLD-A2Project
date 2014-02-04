Public Class LoginScreen

    Private Sub loginbtn_Click(sender As Object, e As EventArgs) Handles loginbtn.Click
        '     If (usernametxt.Text <> My.Settings.username) Then WrongCredentials()
        '     If (passwordtxt.Text <> My.Settings.password) Then WrongCredentials()
        '     If ((usernametxt.Text = My.Settings.username And passwordtxt.Text = My.Settings.password)) Then
        Me.Hide()
        MainmenuForm.Show()
        'End If
    End Sub

    Private Sub clearfieldbtn_Click(sender As Object, e As EventArgs) Handles clearfieldbtn.Click
        usernametxt.Clear() 'Clears the username field
        passwordtxt.Clear() 'Clears the password field
        textboxreset() 'Calls sub textboxrest()
    End Sub

    Private Sub usernametxt_TextChanged(sender As Object, e As EventArgs) Handles usernametxt.TextChanged
        textboxreset() 'Resets the fields
    End Sub

    Private Sub passwordtxt_TextChanged(sender As Object, e As EventArgs) Handles passwordtxt.TextChanged
        textboxreset() 'Resets the fields
    End Sub

    Public Sub WrongCredentials()
        usernametxt.Clear()
        passwordtxt.Clear()
        usernametxt.BackColor = Color.Red
        passwordtxt.BackColor = Color.Red
        MsgBox("Wrong username or password")
    End Sub

    Public Sub textboxreset()
        usernametxt.BackColor = Color.White 'This sets the background colour of the textboxes to white
        passwordtxt.BackColor = Color.White
    End Sub
End Class
