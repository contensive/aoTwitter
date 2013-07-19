VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "processFriendsTimeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'========================================================================
'
'   deprecated -- basic authentication no longer supported
'
'========================================================================
'
'Private Main As ccWeb3.MainClass
Private CSV As Object
Private MainNotUsed As Object
'
' Compatibility method
'
Public Function Main(csvObj As Object, OptionString As String) As String
    Main = Execute(csvObj, Nothing, OptionString, "")
End Function

'
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Dim username As String
    Dim Password As String
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
    Dim ExecPath As String
    '
    Set MainNotUsed = MainObject
    Set CSV = CsvObject
    
    username = CSV.GetSiteProperty("Twitter Username", "", 0)
    Password = CSV.GetSiteProperty("Twitter Password", "", 0)
   
    If (username <> "") And (Password <> "") And kmaEncodeBoolean(CSV.GetSiteProperty("Enable Friends Timeline Process", "0", 0)) Then
        Call Randomize
        Key = Int(100 * Rnd() * 1000)
        OutputFile = CSV.PhysicalFilePath & "Twitter\FriendTimeline\friendtimeline.xml"
        TempFile = CSV.PhysicalFilePath & "Twitter\FriendTimeline\ftl_" & Key & ".xml"
        TempFile2 = CSV.PhysicalFilePath & "Twitter\FriendTimeline\x_ftl_" & Key & ".xml"
        Call CSV.SaveFile(TempFile, "")
        ExecPath = """" & App.Path & "\curl.exe"""
        Cmd = "%comspec% /c " & ExecPath & " -u " & username & ":" & Password & " http://api.twitter.com/1/statuses/friends_timeline.xml > " & TempFile
        'Cmd = "%comspec% /c " & ExecPath & " -u " & username & ":" & Password & " http://twitter.com/statuses/friends_timeline.xml > " & TempFile
        Call kmaShell(Cmd, , True)
        Call objXML.loadXML(CSV.ReadFile(TempFile))
        If objXML.hasChildNodes Then
            For Each Node In objXML.documentElement.childNodes
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
                
                If UCase(Name) <> UCase(username) Then
                    
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
            Output = "<div id=""twitterFollows"">" & vbCrLf & Output & vbCrLf & "</div>"
        End If
        
        Call CSV.DeleteFile(TempFile)
        
        Call CSV.SaveFile(OutputFile, Output)
        
    End If
    
    Exit Function
ErrorTrap:
    Call HandleError("processFriendsTimeline", "Main", Err.Number, Err.Source, Err.Description, True, False)
    End Function

