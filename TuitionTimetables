Imports System.Data.OleDb

Public Class TuitionTimetables

    'Declaring the connection route
   Public connstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DFSArtsDB.accdb"
    'The cursor
    Public conn As New OleDbConnection(connstring)
    Public teacherID As New Integer
    Public studentID As New Integer
    Public SelectIndex As Integer = 0

    Private Sub Homepic_Click(sender As Object, e As EventArgs) Handles Homepic.Click
        Me.Hide()
        MainmenuForm.Show()
    End Sub

    Private Sub TuitionTimetables_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboboxPopulate() 'Calling the ComboboxPopulate sub.
        '    Me.ReportViewer1.RefreshReport()
        SelectIndex = 1
    End Sub

    Public Sub ComboboxPopulate()
        Dim da As New OleDbDataAdapter()
        Dim dt As New DataTable()
        If conn.State = ConnectionState.Closed Then conn.Open()
        da.SelectCommand = New OleDbCommand("select * from tblTeachers", conn)
        da.Fill(dt)
        Tutorcombox.DataSource = dt
        Tutorcombox.DisplayMember = "Surname"
        Tutorcombox.ValueMember = "Surname"
        conn.Close()
    End Sub

    Public Sub InstrumentSelect()
        Dim SelectedTutor As String = Tutorcombox.Text
        Dim da As New OleDbDataAdapter()
        Dim dt As New DataTable()
        If conn.State = ConnectionState.Closed Then conn.Open()
        da.SelectCommand = New OleDbCommand("select * from tblTeachers WHERE Surname = '" & SelectedTutor & "'", conn)
        da.Fill(dt)
        instrumentcombox.DataSource = dt
        instrumentcombox.DisplayMember = "Instrument"
        instrumentcombox.ValueMember = "Instrument"
    End Sub

    Public Sub TeacherIDRetrieve()
        Dim SelectedTutor As String = Tutorcombox.Text
        Dim da As New OleDbDataAdapter()
        Dim dt As New DataTable()
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
        Dim da As New OleDbDataAdapter()
        Dim dt As New DataTable()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Try
            da.SelectCommand = New OleDbCommand("select * from tblAppointments WHERE TeacherID =" & teacherID & "", conn)
            da.Fill(dt)
            Locationcombox.Text = dt.Rows(0).Item(8)
            dt.Columns.Remove("appID")
            dt.Columns.Remove("Present")
            dt.Columns.Remove("Instrument")
            dt.Columns.Remove("Room")
            dt.Columns.Remove("TeacherID")
            testdatagrid.DataSource = dt
        Catch ex As Exception
            MsgBox("There are no appointments in the database for " + Tutorcombox.Text)
        End Try
    End Sub

    Private Sub Tutorcombox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Tutorcombox.SelectedIndexChanged
        If SelectIndex = 0 Then
        Else
            InstrumentSelect()
            TeacherIDRetrieve()
            GetAppoinments()
        End If
    End Sub
End Class
