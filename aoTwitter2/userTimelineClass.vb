Namespace Contensive.Addons.aoTwitter
    '
    Public Class userTimelineClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Dim returnHtml As String = ""
            Dim hint As String = "enter"
            '
            Try
                Dim inS As String = ""
                Dim username As String = CP.Doc.Var("Username")
                Dim followMeCaption As String = CP.Doc.Var("Follow Me Caption")
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
                Dim cacheName As String
                Dim cacheDateLine As String
                Dim cacheDate As Date
                Dim ptr As Integer
                Dim cacheOK As Boolean = False
                Dim useCache As Boolean = False
                Dim cacheEntity As String = ""
                Dim cacheTimeoutSeconds As Integer = 0
                '
                hint &= ",1"
                If count = 0 Then
                    count = 5
                End If
                '
                hint &= ",2"
                '
                hint &= ",3"
                If username = "" Then
                    username = "contensivenews"
                    title = "Contensive News"
                End If
                If title = "" Then
                    title = username
                End If
                If followMeCaption = "" Then
                    followMeCaption = title
                End If
                cacheName = "twitter-" & username
                cacheEntity = CP.Cache.Read(cacheName)
                ptr = cacheEntity.IndexOf(vbCrLf)
                If (ptr = -1) Then
                    cacheEntity = ""
                Else
                    cacheDateLine = cacheEntity.Substring(0, ptr)
                    cacheEntity = cacheEntity.Substring(ptr + 2)
                    If Not IsDate(cacheDateLine) Then
                        cacheEntity = ""
                    Else
                        cacheOK = True
                        cacheDate = CDate(cacheDateLine)
                        If (Date.Now < cacheDate) Then
                            useCache = True
                        End If
                    End If
                End If
                If Not useCache Then
                    '
                    tService.AuthenticateWith(_accessToken, _accessTokenSecret)
                    tOptions.ScreenName = username
                    tOptions.Count = count
                    '
                    Dim tweets As System.Collections.Generic.IEnumerable(Of TweetSharp.TwitterStatus) = tService.ListTweetsOnUserTimeline(tOptions)
                    '
                    If tweets Is Nothing Then
                        '
                        ' exit with best cache we have
                        '
                        useCache = True
                    Else
                        '
                        ' attempt to get new tweets and save new cache
                        '
                        hint &= ",6"
                        hint &= ",7"
                        If title <> "" Then
                            returnHtml += CP.Html.h2(title, , "sidebar-title")
                        End If
                        '
                        hint &= ",8"

                        For Each tweet As TweetSharp.TwitterStatus In tweets
                            hint &= ",9"
                            tweetCopy = tweet.Text
                            numberOfDays = CStr(dToday.Subtract(tweet.CreatedDate).Days)
                            '
                            hint &= ",10"
                            If numberOfDays > 0 Then
                                postedString = CP.Html.p("about " & numberOfDays & " days ago", , "twitterPostedDate")
                            End If
                            '
                            hint &= ",11"
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
                            hint &= ",12"
                            'inS += CP.Html.li(tweetCopy & CP.Html.p(tweet.CreatedDate.ToShortDateString))
                            inS += CP.Html.li(tweetCopy & postedString)
                        Next
                        '
                        hint &= ",13"
                        returnHtml += CP.Html.ul(inS, , , "twitter_update_list")
                        returnHtml += CP.Html.div("<a target=""_blank"" class=""twitter-link"" href=""" & CP.Request.Protocol & "twitter.com/" & username & """ id=""twitter-link"">follow " & followMeCaption & " on Twitter</a>", , , "twitter-link_div")
                        returnHtml = CP.Html.div(returnHtml, , "twitter_div", "twitter_div")
                        cacheTimeoutSeconds = 60 + (60 * Rnd())
                        Call CP.Cache.Save("twitter-" & username, Date.Now.AddSeconds(cacheTimeoutSeconds) & vbCrLf & returnHtml)
                    End If
                End If
                If useCache And cacheOK Then
                    returnHtml = cacheEntity
                End If
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.aoTwitter.userTimelineClass.userTimelineClass, hint=[" & hint & "]")
                Catch errObj As Exception
                End Try
            End Try
            '
            Return returnHtml
        End Function
        '
    End Class
    '
End Namespace
