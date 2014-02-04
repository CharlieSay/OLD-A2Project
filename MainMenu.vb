Imports System.Net
Imports System.Text.RegularExpressions
Public Class MainmenuForm

    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '   TwitterScrape()
        'Makes a integer between 1 and 6
        Dim RandomVal As Integer = CInt(Int((5 * Rnd()) + 1))
        'Creates a webclient that is needed to download the data.
        Dim DFSWebClient As New System.Net.WebClient
        If RandomVal = 1 Then
            'BYTE ARRAY HOLDS THE DATA
            Dim ImageInBytes() As Byte = DFSWebClient.DownloadData("http://davenantperformingarts.org.uk/Content/SlideShowHome/_Images/Bach%20Twitter.jpg")
            'CREATE A MEMORY STREAM USING THE BYTES
            Dim ImageStream As New IO.MemoryStream(ImageInBytes)
            'CREATE A BITMAP FROM THE MEMORY STREAM
            PicturePull.Image = New System.Drawing.Bitmap(ImageStream)
        ElseIf RandomVal = 2 Then
            Dim ImageInBytes() As Byte = DFSWebClient.DownloadData("http://davenantperformingarts.org.uk/Content/SlideShowHome/_Images/1.jpg")
            Dim ImageStream As New IO.MemoryStream(ImageInBytes)
            PicturePull.Image = New System.Drawing.Bitmap(ImageStream)
        ElseIf RandomVal = 3 Then
            Dim ImageInBytes() As Byte = DFSWebClient.DownloadData("http://davenantperformingarts.org.uk/Content/SlideShowHome/_Images/Christmas%20Banner%202013.jpg")
            Dim ImageStream As New IO.MemoryStream(ImageInBytes)
            PicturePull.Image = New System.Drawing.Bitmap(ImageStream)
        ElseIf RandomVal = 4 Then
            Dim ImageInBytes() As Byte = DFSWebClient.DownloadData("http://davenantperformingarts.org.uk/Content/SlideShowHome/_Images/Christmas20121.jpg")
            Dim ImageStream As New IO.MemoryStream(ImageInBytes)
            PicturePull.Image = New System.Drawing.Bitmap(ImageStream)
        Else
            Dim ImageInBytes() As Byte = DFSWebClient.DownloadData("http://davenantperformingarts.org.uk/Content/SlideShowHome/_Images/Picture1.jpg")
            Dim ImageStream As New IO.MemoryStream(ImageInBytes)
            PicturePull.Image = New System.Drawing.Bitmap(ImageStream)
        End If
    End Sub

    Private Sub Tuitionttbtn_Click(sender As Object, e As EventArgs) Handles Tuitionttbtn.Click
        Me.Close()
        TuitionTimetables.Show()
    End Sub

    Private Sub Tutionregistersbtn_Click(sender As Object, e As EventArgs) Handles Tuitionregisterstbn.Click
        Me.Close()
        TuitionRegisters.Show()
    End Sub

    Private Sub Viewpaymentsbtn_Click(sender As Object, e As EventArgs) Handles Viewpaymentsbtn.Click
        Me.Close()
        PaymentViewer.Show()
    End Sub

    Private Sub Recordpaymentsbtn_Click(sender As Object, e As EventArgs) Handles Recordpaymentsbtn.Click
        RecordingPayments.Show()
        Me.Close()
    End Sub

    Private Function GetBetween(ByVal Source As String, ByVal Str1 As String, ByVal Str2 As String, Optional ByVal Index As Integer = 0) As String
        Return Regex.Split(Regex.Split(Source, Str1)(Index + 1), Str2)(0)
    End Function

    Private Function GetBetweenAll(ByVal Source As String, ByVal Str1 As String, ByVal Str2 As String) As String()
        Dim Results, T As New List(Of String) 'Declares collection strings as a list
        T.AddRange(Regex.Split(Source, Str1)) 'Splits the regular expression.
        T.RemoveAt(0) 'Removes certain tags
        For Each I As String In T
            Results.Add(Regex.Split(I, Str2)(0)) 'Adds split strings to results
        Next
        Return Results.ToArray 'Turns results into array and returns it back to the TwitterScrape Sub
    End Function

    'This sub gets the Twitter profile feed of the Performing Arts Department
    'And then converts it to string and runs through various processes to remove certain
    'characters and then presents it.
    Private Sub TwitterScrape()
        'Creating a web request with the url http: http://www.twitter.com/DFSPerfArts
        Dim r As HttpWebRequest = HttpWebRequest.Create("http://www.twitter.com/DFSPerfArts")
        'Provides a container for the incoming response
        Dim re As HttpWebResponse = r.GetResponse()
        'Creates the source string
        Dim src As String = New System.IO.StreamReader(re.GetResponseStream()).ReadToEnd()
        'If there is no source response.
        If (src = Nothing) Then
            MsgBox("Error. Source is null - This means that Twitter is probably down")
        Else
            'Gets the tweets as string collection by using the GetBetweenAll function, which uses javascript to retireve the feed
            Dim tweets As String() = GetBetweenAll(src, "<li class=""js-stream-item stream-item stream-item expanding-stream-item"" data-item-id=""", "</div></div></li>")
            If (tweets.Count > 0) Then
                'Intializing the tweetcount
                Dim tweetcount As Integer = 0
                'Iterates through each tweet.
                For Each tweet As String In tweets '
                    'Increases tweet count
                    tweetcount += 1
                    'Calls getbetween function and gives its variable by value
                    Dim msg As String = GetBetween(tweet, "<p class=""js-tweet-text tweet-text"">", "</p>")
                    'Calls the clearTags function
                    msg = clearTags(msg)
                    'Checks if the tweet is a reply
                    If Mid(msg, 1, 1) = "@" Then
                        Twitterlistbox.Items.Add("Reply : " + msg)
                    Else 'or is a actual tweet
                        Twitterlistbox.Items.Add("Tweet : " + msg)
                    End If
                Next
            Else 'If the Stream got no tweets.
                MsgBox("Error retrieving Twitter feed - Either Twitter is down or their API Has updated")
            End If
        End If
    End Sub


    'This function clears the unnecesary tags and spaces that are on tweets.
    'It will look for apostrophe's , non-breaking spaces and quotation marks.
    'It then replaces them with ASCII Versions in the program
    Private Function clearTags(ByVal s As String)
        If (s.Contains("<") And s.Contains(">")) Then
            Dim toreturn As String = ""
            Dim shouldadd As Boolean = True
            For Each c As Char In s
                If (c = "<") Then shouldadd = False
                If (c = ">") Then shouldadd = True
                If (Not c = "<" And Not c = ">" And shouldadd) Then
                    toreturn &= c
                End If
            Next
            If (toreturn.Contains("&#39;")) Then
                toreturn = toreturn.Replace("&#39;", "'")
            End If
            If (toreturn.Contains("&nbsp;")) Then
                toreturn = toreturn.Replace("&nbsp;", " ")
            End If
            If (toreturn.Contains("&quot;")) Then
                toreturn = toreturn.Replace("&quot;", """")
            End If
            Return toreturn
        Else
            Dim s2 As String = ""
            If (s2.Contains("&#39;")) Then
                s2 = s2.Replace("&#39;", "'")
            End If
            If (s2.Contains("&nbsp;")) Then
                s2 = s2.Replace("&nbsp;", " ")
            End If
            If (s2.Contains("&quot;")) Then
                s2 = s2.Replace("&quot;", """")
            End If
            Return s2
        End If
    End Function
End Class
