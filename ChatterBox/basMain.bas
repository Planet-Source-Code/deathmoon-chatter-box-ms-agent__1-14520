Attribute VB_Name = "basMain"
Option Explicit

Public cAgent As IAgentCtlCharacterEx
Public sAgentPath As String
Public CurrentTime As Double
Public sNickName As String
Public sRemoteHost As String
Public sPort As String
Public bChar As Byte
Private booSecondTime As Boolean

Private Const sAppName As String = "Chatter Box"
Private Const sSection As String = "User Options"

Sub Main()
    ' LOAD HELP FILE
    App.HelpFile = ""
    '
    OSVersion_And_AgentFiles
    '
    frmSplash.Show
End Sub

Sub OSVersion_And_AgentFiles()
    Dim sOSVer As String
    sOSVer = basWhatOS.GetVersion   ' Returns the O/S
    If sOSVer = "Windows NT 4.0" Then
        sAgentPath = "C:\WINNT\msagent\chars"
    ElseIf sOSVer = "Windows 95" Then
        sAgentPath = "c:\windows\msagent\chars"
    ElseIf sOSVer = "Windows 98" Then
        sAgentPath = "c:\windows\msagent\chars"
    End If
End Sub

Sub Wait(interval)
    CurrentTime = Timer
    Do While Timer - CurrentTime < Val(interval)
        DoEvents
    Loop
End Sub

Sub FillCombo(ctl As ComboBox)
    With ctl
        .AddItem "Genie"
        .AddItem "Merlin"
        .AddItem "Robby"
        .AddItem "Peedy"
    End With
End Sub

Sub MovementsForCharacters(sChar As String, _
    aAgt As IAgentCtlCharacterEx)
    ' THIS DETERMINES WHAT CHARACTER THE USER IS USING
    ' AND THEN WILL OPEN THE CORRECT MODULE FOR THAT
    ' CHARACTER.  THIS WILL ENABLE THAT CHARACTER TO
    ' PERFORM DIFFERENT FEATS, ETC.
    
    If sChar = "Merlin" Then
        bChar = 1
    ElseIf sChar = "Robby" Then
        bChar = 2
    ElseIf sChar = "Genie" Then
        bChar = 3
    ElseIf sChar = "Peddy" Then
        bChar = 4
    End If
    
    If booSecondTime = False Then
        If bChar = 1 Then
            basMerlin.m_nMerlin cAgent, 1, "LoadAgent"
        ElseIf bChar = 2 Then
        
        ElseIf bChar = 3 Then
        
        ElseIf bChar = 4 Then
        
        End If
    End If
    
    booSecondTime = True
End Sub

Sub SaveUsrSettings(sNickName As String, _
    sRemoteHost As String, sPort As String)
    
    SaveSetting sAppName, sSection, _
        "RemoteHost", sRemoteHost
    SaveSetting sAppName, sSection, _
        "NickName", sNickName
    SaveSetting sAppName, sSection, _
        "PortNumber", sPort
End Sub

Sub GetUsrSettings()
    sRemoteHost = GetSetting(sAppName, sSection, "RemoteHost")
    sNickName = GetSetting(sAppName, sSection, "NickName")
    sPort = GetSetting(sAppName, sSection, "PortNumber")
End Sub
