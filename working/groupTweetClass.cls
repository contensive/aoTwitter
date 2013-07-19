VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "groupTweetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameMessage = "msg"
Const FormButtonSubmit = " Post Tweet "
Const cr = vbCrLf & vbTab
Const cr2 = vbCrLf & vbTab & vbTab
'
Private ErrorString As String
Private username As String
Private Password As String
Private ExecPath As String
Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String

    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    
    username = Main.GetSiteProperty("Twitter Username", "")
    Password = Main.GetSiteProperty("Twitter Password", "")
    
    ExecPath = """" & App.Path & "\curl.exe"""
    
    'Call AppendLog("groupTweets" & CDbl(Date) & ".txt", "groupTweetClass[groupTweetClass] - ExecPath: " & ExecPath)
    
    Call ProcessForm
    
    Exit Sub
ErrorTrap:
    Call HandleError("groupTweetClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim username As String
    '
    username = Main.GetSiteProperty("Twitter Username", "")
    '
    Stream = Main.GetFormStart
    Stream = Stream & cr & "<div id=""groupTweetForm"">"
    Stream = Stream & cr & "<h2 class=""headline"">Post a New Tweet</h2>"
    Stream = Stream & cr & "<p>Use this tool to post a tweet to the account configured in Twitter Configuration under Settings on the Admin Navigator.<p>"
    If username = "" Then
        Stream = Stream & cr & "<p>The username configured is currently blank. You must set a username and password before attempting a tweet.</p>"
        Stream = Stream & cr & "<p><a href=""" & Main.GetSiteProperty("adminUrl") & "?addonguid=%7B0C107FCD-B5CE-4B16-91EC-1103C5896FE0%7D"">Configure Your Twitter Account</a></p>"
    Else
        Stream = Stream & cr & "<p>The username configured is <b>" & username & "</b>.<p>"
        Stream = Stream & cr2 & "<div>" & Main.GetFormInputText(RequestNameMessage, "", 6, 45) & "</div>"
        Stream = Stream & cr2 & "<div>" & Main.GetFormButton(FormButtonSubmit) & "</div>"
        Stream = Stream & GetUserTweets()
    End If
    Stream = Stream & cr & "</div>"
    Stream = Stream & Main.GetFormEnd
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("groupTweetClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
End Function

Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    
    Dim Message As String
    Dim Cmd As String
    Dim objShell As New WshShell
    
    If Main.GetStreamText("button") = FormButtonSubmit Then
        
        Message = Main.GetStreamText(RequestNameMessage)
        
        Cmd = ExecPath & " -u " & username & ":" & Password & " -d status=""" & Message & """ http://api.twitter.com/1/statuses/update.xml"
        'Cmd = ExecPath & " -u " & username & ":" & Password & " -d status=""" & Message & """ http://www.twitter.com/statuses/update.xml"
        
        'Call AppendLog("groupTweets" & CDbl(Date) & ".txt", "groupTweetClass[ProcessForm] - CMD: " & CMD)
        
        Call objShell.Run(Cmd, 0, True)
        
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("groupTweetClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Function GetUserTweets() As String
    On Error GoTo ErrorTrap
    
    Dim Cmd As String
    Dim OutputFile As String
    Dim TempFile As String
    Dim TempFile2 As String
    Dim Key As String
    Dim objXML As New MSXML2.DOMDocument
    Dim Node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim UserNode As IXMLDOMNode
    Dim Output As String
    Dim ThisOutput As String
    
    Dim Following As Boolean
    Dim Name As String
    Dim URL As String
    Dim Created As String
    Dim Message As String
    Dim Image As String
    
    Dim PosStart As Long
    Dim PosEnd As Long
    Dim FindLink As String
    Dim MessageEnd As Long
    Dim Pointer As Long
    
    Dim TimePosStart As Long
    Dim TimePosEnd As Long
    Dim Time As String
    Dim TimeDif As String
    Dim TimePart As String
        
    Call Randomize
    
    Key = Int(100 * Rnd() * 1000)
    
    OutputFile = Main.PhysicalFilePath & "Twitter\GroupTweets\usertimeline.xml"
    TempFile = Main.PhysicalFilePath & "Twitter\GroupTweets\usertimeline_" & Key & ".xml"
    TempFile2 = Main.PhysicalFilePath & "Twitter\GroupTweets\x_usertimeline_" & Key & ".xml"
    
    Call Main.SaveFile(TempFile, "")
    
    Cmd = "%comspec% /c " & ExecPath & " -u " & username & ":" & Password & " http://api.twitter.com/1/statuses/user_timeline.xml > " & TempFile
    'Cmd = "%comspec% /c " & ExecPath & " -u " & username & ":" & Password & " http://twitter.com/statuses/user_timeline.xml > " & TempFile
    
    'Call AppendLog("groupTweets" & CDbl(Date) & ".txt", "groupTweetClass[GetUserTweets] - CMD: " & Cmd)
    
    Call kmaShell(Cmd, , True)

    Call objXML.loadXML(Main.ReadFile(TempFile))
    If objXML.hasChildNodes Then
        For Each Node In objXML.documentElement.childNodes
            Message = ""
            If Node.hasChildNodes Then
                For Each InnerNode In Node.childNodes
                    Select Case InnerNode.nodeName
                        Case "created_at"
                            Created = InnerNode.Text
                        Case "text"
                            Message = InnerNode.Text & vbCrLf
                        Case "user"
                            If InnerNode.hasChildNodes Then
                                For Each UserNode In InnerNode.childNodes
                                    Select Case UserNode.nodeName
                                        Case "following"
                                            Following = kmaEncodeBoolean(UserNode.Text)
                                        Case "screen_name"
                                            Name = UserNode.Text
                                        Case "url"
                                            URL = UserNode.Text
                                        Case "profile_image_url"
                                            Image = UserNode.Text
                                    End Select
                                Next
                            End If
                    End Select
                Next
            End If
                
            If URL <> "" Then
                If (InStr(1, URL, "http://", vbTextCompare) = 0) And (InStr(1, URL, "http//", vbTextCompare) = 0) Then
                    URL = "http://" & URL
                End If
                Name = "<a target=""_blank"" href=""" & URL & """>" & Name & "</a>"
            End If
            
            If Created <> "" Then
                If InStr(1, Created, "+", vbTextCompare) Then
                    TimePosEnd = InStr(1, Created, "+", vbTextCompare)
                    TimePosEnd = InStr(TimePosEnd, Created, " ", vbTextCompare)
                    TimePosStart = InStr(1, Created, ":", vbTextCompare)
                    TimePosStart = InStrRev(Created, " ", TimePosStart, vbTextCompare)
                    Time = Mid(Created, TimePosStart + 1, TimePosEnd - TimePosStart)
                    Created = Replace(Created, Time, "", , , vbTextCompare)
                End If
            End If
            '
            If Time <> "" Then
                Time = Trim(Time)
                TimePosStart = InStr(1, Time, " ", vbTextCompare)
                TimeDif = Right(Time, Len(Time) - TimePosStart)
                Time = Replace(Time, TimeDif, "", , , vbTextCompare)
                TimePosStart = InStr(1, Time, ":", vbTextCompare) - 1
                TimePart = Left(Time, TimePosStart)
                Time = Mid(Time, TimePosStart + 1, Len(Time))
                If (TimePart > 12) Or (TimePart = "00") Then
                    TimePart = TimePart - 12
                    If (TimePart = "00") Then
                        TimePart = 12
                    End If
                    If (TimePart = 0) Then
                        TimePart = 12
                    End If
                    Time = TimePart & Time & " PM"
                Else
                    Time = TimePart & Time & " AM"
                End If
            End If
            '
            If Message <> "" Then
                MessageEnd = Len(Message)
                If InStr(1, Message, "http://", vbTextCompare) <> 0 Then
                    Do While Pointer < MessageEnd
                        PosStart = InStr(1, Message, "http://", vbTextCompare)
                        If PosStart = 0 Then
                            PosStart = 1
                        End If
                        PosEnd = InStr(PosStart, Message, " ", vbTextCompare)
                        If PosEnd = 0 Then
                            PosEnd = MessageEnd
                        End If
                        Pointer = PosEnd
                        FindLink = Mid(Message, PosStart, PosEnd - PosStart)
                        Message = Replace(Message, FindLink, "<a target=""_blank"" href=""" & FindLink & """>" & FindLink & "</a>", , , vbTextCompare)
                    Loop
                    Pointer = 0
                End If
                Output = Output & "<table cellspacing=0 class=""twitter_message"">" & vbCrLf
                Output = Output & "<tr>" & vbCrLf
                Output = Output & "<td valign=""top"" class=""profileImage""><img src=""" & Image & """></td>" & vbCrLf
                Output = Output & "<td valign=""top"" width=""100%"" class=""followerMessage"">" & vbCrLf
                Output = Output & "<div class=""headline"">" & Name & "</div>" & vbCrLf
                Output = Output & "<div class=""created"">" & Created & " - " & Time & "</div>" & vbCrLf
                Output = Output & "<div class=""message"">" & Message & "</div>" & vbCrLf
                Output = Output & "</td>" & vbCrLf
                Output = Output & "</tr>" & vbCrLf
                Output = Output & "</table>" & vbCrLf
            End If
        Next
    End If
    
    If Output <> "" Then
        Output = "<div id=""twitterUserTimeline"">" & vbCrLf & Output & vbCrLf & "</div>"
    End If
    
    Call Main.DeleteFile(TempFile)
    
    GetUserTweets = Output
    
    Exit Function
ErrorTrap:
    Call HandleError("groupTweetClass", "GetUserTweets", Err.Number, Err.Source, Err.Description, True, False)
End Function
