Imports System.Data.OleDb

Public Class TuitionRegisters


    'Declaring the connection route
    Public connstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DFSArtsDB.accdb"
    'The cursor
    Public conn As New OleDbConnection(connstring)
    Public teacherID As New Integer
    Public studentID As New Integer
    Public SelectIndex As Integer = 0
    Public DisplayTable As DataTable
    Public Finalda As New OleDbDataAdapter
    Public Finalds As New DataSet

    Private Sub Homepic_Click(sender As Object, e As EventArgs) Handles Homepic.Click
        Me.Close()
        MainmenuForm.Show()
    End Sub

    Private Sub TuitionTimetables_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboboxPopulate()
        SelectIndex = 1
    End Sub

    Public Sub ComboboxPopulate()
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable()
        If conn.State = ConnectionState.Closed Then conn.Open()
        da.SelectCommand = New OleDbCommand("select * from tblTeachers", conn)
        da.Fill(dt)
        Tutorcombox.DataSource = dt
        Tutorcombox.DisplayMember = "Surname"
        Tutorcombox.ValueMember = "Surname"
        conn.Close()
    End Sub

    Public Sub TeacherIDRetrieve()
        Dim SelectedTutor As String = Tutorcombox.Text
        Dim dt As New DataTable()
        Dim da As New OleDbDataAdapter()
        If conn.State = ConnectionState.Closed Then conn.Open()
        da.SelectCommand = New OleDbCommand("select * from tblTeachers WHERE Surname = '" & SelectedTutor & "'", conn)
        da.Fill(dt)
        For Each row As DataRow In dt.Rows
            teacherID = row.Item(0)
        Next
        MsgBox(teacherID)
    End Sub

    Public Sub GetAppoinments()
        Dim SearchID As Integer = teacherID
        Dim NewStudentID As Integer = studentID
        Dim DisplayTable As New DataTable()
        DisplayTable.Clear()
        Dim da As New OleDbDataAdapter()
        Dim sqlquery As String = ("select * from tblAppointments WHERE TeacherID =" & teacherID & "")
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Try
            da.SelectCommand = New OleDbCommand(sqlquery, conn)
            da.Fill(Finalds, "Display")
            DisplayTable = Finalds.Tables("Display")
            DisplayTable.Columns.Remove("Instrument")
            DisplayTable.Columns.Remove("Room")
            DisplayTable.Columns.Remove("TeacherID")
            Registersgridview.DataSource = DisplayTable
            conn.Close()
        Catch ex As Exception
            MsgBox("There are no appointments in the database for " + Tutorcombox.Text)
        End Try
    End Sub

    Private Sub Tutorcombox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Tutorcombox.SelectedIndexChanged
        If SelectIndex = 0 Then
        Else
            TeacherIDRetrieve()
            GetAppoinments()
        End If
    End Sub

    Private Sub Discardchangesbtn_Click(sender As Object, e As EventArgs) Handles Discardchangesbtn.Click
        '  Registersgridview.DataSource = Nothing
        MsgBox("Discarded Changes!")
    End Sub

    Private Sub AcceptChangeBtn_Click(sender As Object, e As EventArgs) Handles AcceptChangeBtn.Click
        Dim dt As DataTable = New DataTable("SendTable")
        Dim row As DataRow
        dt.Columns.Add("appID", Type.GetType("System.Int32"))
        dt.Columns.Add("Present", Type.GetType("System.Boolean"))
        For i = 0 To Registersgridview.Rows.Count - 1
            Dim appID As Integer = Registersgridview.Rows(i).Cells(0).Value
            Dim present As Boolean = Registersgridview.Rows(i).Cells(4).Value
            row = dt.Rows.Add
            row.Item("appID") = appID
            row.Item("Present") = present
        Next
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim sqlquery As String = "UPDATE tblAppointments SET Present = @Present WHERE appID = @appID"
        Dim sqlcommand As New OleDbCommand
        For Each newrow As DataRow In dt.Rows
            With sqlcommand
                .CommandText = sqlquery
                MsgBox(newrow.Item(1))
                MsgBox(newrow.Item(0))
                .Parameters.AddWithValue("@Present", newrow.Item(1))
                .Parameters.AddWithValue("@appID", newrow.Item(0))
                .Connection = conn
                .ExecuteNonQuery()
            End With
        Next
        conn.Close()
        Registersgridview.DataSource = Nothing
        MsgBox("Completed")
        dt.Clear()
    End Sub
End Class
