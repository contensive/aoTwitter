VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "followingTimelineClass"
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
    Call HandleError("followingTimelineClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim username As String
    Dim Password As String
    Dim copy As String
    '
    username = CSV.GetSiteProperty("Twitter Username", "", 0)
    Password = CSV.GetSiteProperty("Twitter Password", "", 0)
    '
    If (username <> "") And (Password <> "") And kmaEncodeBoolean(CSV.GetSiteProperty("Enable Friends Timeline Process", "0", 0)) Then
        GetContent = Main.ReadFile(Main.PhysicalFilePath & "Twitter\FriendTimeline\friendtimeline.xml")
    Else
        copy = "" _
            & " This addon creates a list of tweets that you follow on your Twitter account.</p>" _
            & "<p>Please setup your Twitter username/password and enable 'Friends Timeline Processing' in the Twitter Configuration under settings.</p>" _
            & "<p><a href=""" & Main.GetSiteProperty("adminurl") & "?addonguid=%7B0C107FCD-B5CE-4B16-91EC-1103C5896FE0%7D"">Configure Your Twitter Account</a>"
        GetContent = Main.GetAdminHintWrapper(copy)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("followingTimelineClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
End Function


