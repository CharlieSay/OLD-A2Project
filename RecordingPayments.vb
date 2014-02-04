Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class RecordingPayments

    'Declaring the connection route
    Public connstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DFSArtsDB.accdb"
    'The cursor
    Public conn As New OleDbConnection(connstring)

    Dim SForename As String
    Dim SSurname As String
    Dim SForm As String
    Dim SPaid As Integer

    Private Sub Clearfieldbtn_Click(sender As Object, e As EventArgs) Handles Clearfieldbtn.Click
        Forenametxt.Clear()
        Surnametxt.Clear()
        Formtxt.Clear()
        Paidtxt.Clear()
    End Sub

    Private Sub Homepic_Click(sender As Object, e As EventArgs) Handles Homepic.Click
        Me.Hide()
        MainmenuForm.Show()
    End Sub

    Private Sub Addrecordbtn_Click(sender As Object, e As EventArgs) Handles Addrecordbtn.Click
        If Integer.TryParse(Paidtxt.Text, SPaid) Then
            'Checks the data type of Paidtxt against a integer variable.        
            If Paidtxt.Text > 200 Then
                MsgBox("Payment field is to large!")
            ElseIf Paidtxt.Text < 10 Then
                MsgBox("Payment field is too small")
            Else
                RecordingPayments()
            End If
        Else
            MsgBox("Payment field isnt an integer!")
            Paidtxt.BackColor = Color.Red
        End If
        If Forenametxt.Text.Length < 1 Or Forenametxt.Text.Length > 30 Then
            MsgBox("Forename field has an incorrect length!")
            Forenametxt.BackColor = Color.Red
        End If
        If Surnametxt.Text.Length < 1 Or Surnametxt.Text.Length > 30 Then
            MsgBox("Forename field has an incorrect length!")
            Surnametxt.BackColor = Color.Red
        End If
        If Formtxt.Text.Length > 4 Or Formtxt.Text.Length < 2 Then
            MsgBox("Form field has an incorrect length")
            Formtxt.BackColor = Color.Red
        End If
        If Paidtxt.Text = "" Then
            MsgBox("Paid field is empty!")
            Paidtxt.BackColor = Color.Red
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Paidtxt.Text = "300"
        Forenametxt.Text = "Charlie"
        Surnametxt.Text = "Say"
        Formtxt.Text = "SF7"
    End Sub

    Public Sub RecordingPayments()
        'SQL Query command for ADDING
        Dim sqlquery As String = "INSERT INTO tblPayments(StudentID, Payment, PaymentDate) SELECT StudentID,  @SPaid, @todaysdate FROM tblStudents WHERE Forename = @Forename AND Surname = @Surname"
        'Creating the command itself.
        Dim sqlcommand As New OleDbCommand
        Try
            With sqlcommand
                'Telling what query to execute.
                .CommandText = sqlquery
                '  Paramaters to add with values.
                .Parameters.AddWithValue("@SPaid", Paidtxt.Text)
                .Parameters.AddWithValue("@todaysdate", Today.Date)
                .Parameters.AddWithValue("@Forename", Forenametxt.Text)
                .Parameters.AddWithValue("@Surname", Surnametxt.Text)
                '   Selecting the connection
                .Connection = conn
                '  Executing the query
                .ExecuteNonQuery()
                MsgBox("query executed, closing connection")
            End With
        Catch sqlex As SqlException
            MsgBox("SQL Error")
        Catch ex As Exception
            MsgBox("Error - Either the connection has broken to the DB or that student doesn't exist in the DB!")
        End Try
        Forenametxt.Clear()
        Surnametxt.Clear()
        Formtxt.Clear()
        Paidtxt.Clear()
        conn.Close()
    End Sub

    Private Sub Forenametxt_TextChanged(sender As Object, e As EventArgs) Handles Forenametxt.TextChanged
        Forenametxt.BackColor = Color.White
    End Sub

    Private Sub Surnametxt_TextChanged(sender As Object, e As EventArgs) Handles Surnametxt.TextChanged
        Surnametxt.BackColor = Color.White
    End Sub

    Private Sub Formtxt_TextChanged(sender As Object, e As EventArgs) Handles Formtxt.TextChanged
        Formtxt.BackColor = Color.White
    End Sub

    Private Sub Paidtxt_TextChanged(sender As Object, e As EventArgs) Handles Paidtxt.TextChanged
        Paidtxt.BackColor = Color.White
    End Sub
End Class
