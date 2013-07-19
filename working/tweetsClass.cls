VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tweetsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ErrorString As String

Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String

    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    
    Exit Sub
ErrorTrap:
    Call HandleError("tweetsClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim Title As String
    Dim Count As Long
    Dim username As String
    Dim copy As String
    
    username = Main.GetAggrOption("Username", OptionString)
    Title = Main.GetAggrOption("Title", OptionString)
    Count = KmaEncodeNumber(Main.GetAggrOption("Tweets to display", OptionString))
    
    If Count = 0 Then
        Count = 5
    End If
    
    If username <> "" Then
        Stream = "<div class=""twitter_div"" id=""twitter_div"">"
        If Title <> "" Then
            Stream = Stream & "<h2 class=""sidebar-title"">" & Title & "</h2>"
        End If
        Stream = Stream & "<ul id=""twitter_update_list""></ul>"
        Stream = Stream & "<div id=""twitter-link_div"">"
        Stream = Stream & "<a class=""twitter-link"" href=""" & Main.ServerProtocol & "twitter.com/" & username & """ id=""twitter-link"">follow me on Twitter</a>"
        Stream = Stream & "</div>"
        Stream = Stream & "</div>"
        Stream = Stream & "<script type=""text/javascript"" src=""" & Main.ServerProtocol & "twitter.com/javascripts/blogger.js""></script>"
        Stream = Stream & "<script type=""text/javascript"" src=""" & Main.ServerProtocol & "api.twitter.com/1/statuses/user_timeline.json?callback=twitterCallback2&amp;screen_name=" & username & "&amp;count=" & Count & """></script>"
        'Stream = Stream & "<script type=""text/javascript"" src=""" & Main.ServerProtocol & "twitter.com/statuses/user_timeline/" & username & ".json?callback=twitterCallback2&amp;count=" & Count & """></script>"
    Else
        copy = "" _
            & " This addon creates a list of tweets from a specific Twitter account.</p>" _
            & "<p>Please turn on advanced edit and click the addon options icon to set the Twitter username of the account you want to display.</p>" _
            & ""
        Stream = Main.GetAdminHintWrapper(copy)
    End If
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("tweetsClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
End Function

