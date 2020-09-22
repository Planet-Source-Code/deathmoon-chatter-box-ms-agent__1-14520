Attribute VB_Name = "basMerlin"
Option Explicit

' THIS MODULE WILL MAKE THE CHARACTER 'MERLIN' DO
' RANDOM ACTS AS WELL AS MEANINGFUL ACTS.
Public Sub m_nMerlin(aAgt As IAgentCtlCharacterEx, _
    sPlayWhat As Integer, _
    Optional sEvent As String)
    '
    If sPlayWhat > 0 Then
        Select Case sEvent
            Case Is = "SentMsg"
                aAgt.Play ("Pleased")
            Case Is = "LoadAgent"
                aAgt.Play ("greet")
        End Select
        sPlayWhat = 0
    ElseIf sPlayWhat = 0 Then
        aAgt.Play ("read")
    End If
    '
End Sub
