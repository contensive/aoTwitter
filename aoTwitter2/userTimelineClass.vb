Namespace Contensive.Addons.aoTwitter
    '
    Public Class userTimelineClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Dim s As String = ""
            '
            Try
                Dim inS As String = ""
                Dim username As String = CP.Doc.Var("Username")
                Dim title As String = CP.Doc.Var("Title")
                Dim count = CP.Utils.EncodeNumber(CP.Doc.Var("Tweets to display"))
                '
                Dim _consumerKey As String = "rbEicL4D1H0Lc0PwL0oAw"
                Dim _consumerSecret As String = "jKRySGRSp4Qw8bmM2Y2bHLmHEOw5tQ5JqkuCLOdg8T0"
                Dim _accessToken As String = "50011820-sIhp962ouAesENHqDbHVReLuuyAziAh3hNHbGjK78"
                Dim _accessTokenSecret As String = "Q8MnjQlqPh379CNDENzjIYVJI8eqflW32h1slrLEp8"
                Dim tOptions As New TweetSharp.ListTweetsOnUserTimelineOptions
                Dim tService As New TweetSharp.TwitterService(_consumerKey, _consumerSecret)
                '
                Dim tweetCopy As String = ""
                Dim linkStart As Integer = 0
                Dim linkEnd As Integer = 0
                Dim link As String = ""
                Dim numberOfDays As Integer = 0
                Dim dToday As Date = Date.Today
                Dim postedString As String = ""
                '
                If count = 0 Then
                    count = 5
                End If
                '
                tService.AuthenticateWith(_accessToken, _accessTokenSecret)
                '
                If username = "" Then
                    username = "contensivenews"
                    title = "Contensive News"
                End If
                '
                tOptions.ScreenName = username
                tOptions.Count = count
                '
                Dim tweets As System.Collections.Generic.IEnumerable(Of TweetSharp.TwitterStatus) = tService.ListTweetsOnUserTimeline(tOptions)
                '
                If username <> "" Then
                    If title <> "" Then
                        s += CP.Html.h2(title, , "sidebar-title")
                    End If
                    '
                    For Each tweet As TweetSharp.TwitterStatus In tweets
                        tweetCopy = tweet.Text
                        numberOfDays = CStr(dToday.Subtract(tweet.CreatedDate).Days)
                        '
                        If numberOfDays > 0 Then
                            postedString = CP.Html.p("about " & numberOfDays & " days ago", , "twitterPostedDate")
                        End If
                        '
                        If tweetCopy.Contains("http") Then
                            linkStart = tweetCopy.IndexOf("http")
                            linkEnd = tweetCopy.IndexOf(" ", linkStart)
                            '
                            If linkEnd = -1 Then
                                linkEnd = tweetCopy.Length
                            End If
                            '
                            link = tweetCopy.Substring(linkStart, linkEnd - linkStart)
                            '
                            tweetCopy = tweetCopy.Replace(link, "<a target=""_blank"" href=""" & link & """>" & link & "</a>")
                        End If
                        '
                        'inS += CP.Html.li(tweetCopy & CP.Html.p(tweet.CreatedDate.ToShortDateString))
                        inS += CP.Html.li(tweetCopy & postedString)
                    Next
                    '
                    s += CP.Html.ul(inS, , , "twitter_update_list")
                    s += CP.Html.div("<a class=""twitter-link"" href=""" & CP.Request.Protocol & "twitter.com/" & username & """ id=""twitter-link"">follow me on Twitter</a>", , , "twitter-link_div")
                    s = CP.Html.div(s, , "twitter_div", "twitter_div")
                Else
                    s = "" _
                        & " This addon creates a list of tweets from a specific Twitter account.</p>" _
                        & "<p>Please turn on advanced edit and click the addon options icon to set the Twitter username of the account you want to display.</p>" _
                        & ""
                    '
                    s = CP.Html.div(s, , "ccEditWrapperCaption")
                End If
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.aoTwitter.userTimelineClass.userTimelineClass")
                Catch errObj As Exception
                End Try
            End Try
            '
            Return s
        End Function
        '
    End Class
    '
End Namespace
