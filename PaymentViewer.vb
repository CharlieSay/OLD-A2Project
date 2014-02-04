Imports System.Data.OleDb

Public Class PaymentViewer

    'Declaring the connection route
    Public connstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DFSArtsDB.accdb"
    'The cursor
    Public conn As New OleDbConnection(connstring)

    Public FirstDayBefore As DateTime
    Public Q1 As DateTime
    Public Q2 As DateTime
    Public Q3 As DateTime
    Public Q4 As DateTime
    Public FirstDayAfter As DateTime


    Private Sub Homepic_Click(sender As Object, e As EventArgs) Handles Homepic.Click
        Me.Close()
        MainmenuForm.Show()
    End Sub

    Private Sub SearchBtn_Click(sender As Object, e As EventArgs) Handles SearchBtn.Click
        YearFilter()
        Dim Quarter As String = Quartercombox.Text
        Select Case Quarter
            Case "Q1"
                Dim da As New OleDbDataAdapter()
                Dim ds As New DataSet()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                da.SelectCommand = New OleDbCommand("select * from tblPayments", conn)
                da.Fill(ds, "Payments")
                Dim dt As DataTable = ds.Tables("Payments")
                Dim DisplayTable As DataTable = dt.Clone
                For Each row As DataRow In dt.Rows
                    Dim rowdate As DateTime = row.Item(3)
                    If (rowdate < Q1 And rowdate > FirstDayBefore) Or (rowdate = Q1) Then
                        DisplayTable.ImportRow(row)
                    Else
                    End If
                Next
                With Paymentgridview
                    .AutoGenerateColumns = True
                    .DataSource = DisplayTable
                End With
            Case "Q2"
                Dim da As New OleDbDataAdapter()
                Dim ds As New DataSet()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                da.SelectCommand = New OleDbCommand("select * from tblPayments", conn)
                da.Fill(ds, "Payments")
                Dim dt As DataTable = ds.Tables("Payments")
                Dim DisplayTable As DataTable = dt.Clone
                For Each row As DataRow In dt.Rows
                    Dim rowdate As DateTime = row.Item(3)
                    If (rowdate < Q2 And rowdate > Q1) Or (rowdate = Q2) Then
                        DisplayTable.ImportRow(row)
                    Else
                    End If
                Next
                With Paymentgridview
                    .AutoGenerateColumns = True
                    .DataSource = DisplayTable
                End With
            Case "Q3"
                Dim da As New OleDbDataAdapter()
                Dim ds As New DataSet()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                da.SelectCommand = New OleDbCommand("select * from tblPayments", conn)
                da.Fill(ds, "Payments")
                Dim dt As DataTable = ds.Tables("Payments")
                Dim DisplayTable As DataTable = dt.Clone
                For Each row As DataRow In dt.Rows
                    Dim rowdate As DateTime = row.Item(3)
                    If (rowdate < Q3 And rowdate > Q2) Or (rowdate = Q3) Then
                        DisplayTable.ImportRow(row)
                    Else
                    End If
                Next
                With Paymentgridview
                    .AutoGenerateColumns = True
                    .DataSource = DisplayTable
                End With
            Case "Q4"
                Dim da As New OleDbDataAdapter()
                Dim ds As New DataSet()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                da.SelectCommand = New OleDbCommand("select * from tblPayments", conn)
                da.Fill(ds, "Payments")
                Dim dt As DataTable = ds.Tables("Payments")
                Dim DisplayTable As DataTable = dt.Clone
                For Each row As DataRow In dt.Rows
                    Dim rowdate As DateTime = row.Item(3)
                    If (rowdate > Q3 And rowdate < FirstDayAfter) Or (rowdate = Q4) Then
                        DisplayTable.ImportRow(row)
                    Else
                    End If
                Next
                With Paymentgridview
                    .AutoGenerateColumns = True
                    .DataSource = DisplayTable
                End With
        End Select
    End Sub

    Public Sub YearFilter()
        Dim Year As String = Yeardropdown.Text
        Select Case Year
            Case "2013"
                FirstDayBefore = New DateTime(2012, 12, 31)
                FirstDayAfter = New DateTime(2014, 1, 1)
                Q1 = New DateTime(2013, 3, 31)
                Q2 = New DateTime(2013, 6, 30)
                Q3 = New DateTime(2013, 9, 30)
                Q4 = New DateTime(2013, 12, 31)
            Case "2014"
                FirstDayBefore = New DateTime(2013, 12, 31)
                FirstDayAfter = New DateTime(2015, 1, 1)
                Q1 = New DateTime(2014, 3, 31)
                Q2 = New DateTime(2014, 6, 30)
                Q3 = New DateTime(2014, 9, 30)
                Q4 = New DateTime(2014, 12, 31)
            Case "2015"
                FirstDayBefore = New DateTime(2014, 12, 31)
                FirstDayAfter = New DateTime(2016, 1, 1)
                Q1 = New DateTime(2015, 3, 31)
                Q2 = New DateTime(2015, 6, 30)
                Q3 = New DateTime(2015, 9, 30)
                Q4 = New DateTime(2015, 12, 31)
            Case "2016"
                FirstDayBefore = New DateTime(2015, 12, 31)
                FirstDayAfter = New DateTime(2017, 1, 1)
                Q1 = New DateTime(2016, 3, 31)
                Q2 = New DateTime(2016, 6, 30)
                Q3 = New DateTime(2016, 9, 30)
                Q4 = New DateTime(2016, 12, 31)
            Case "2017"
                FirstDayBefore = New DateTime(2016, 12, 31)
                FirstDayAfter = New DateTime(2018, 1, 1)
                Q1 = New DateTime(2017, 3, 31)
                Q2 = New DateTime(2017, 6, 30)
                Q3 = New DateTime(2017, 9, 30)
                Q4 = New DateTime(2017, 12, 31)
        End Select
    End Sub
End Class
