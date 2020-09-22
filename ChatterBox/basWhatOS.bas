Attribute VB_Name = "basWhatOS"
Option Explicit

' Determine which 32-bit Windows version is running.
' Actual code from Microsoft site. KB# Q189249

Public Declare Function GetVersionExA Lib "kernel32" _
    (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Function GetVersion() As String
    Dim OSInfo As OSVERSIONINFO
    Dim retvalue As Integer
    OSInfo.dwOSVersionInfoSize = 148
    OSInfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(OSInfo)
    With OSInfo
        Select Case .dwPlatformId
            Case 1
                If .dwMinorVersion = 0 Then
                    GetVersion = "Windows 95"
                ElseIf .dwMinorVersion = 10 Then
                    GetVersion = "Windows 98"
                End If
            Case 2
                If .dwMajorVersion = 3 Then
                    GetVersion = "Windows NT 3.51"
                ElseIf .dwMajorVersion = 4 Then
                    GetVersion = "Windows NT 4.0"
                End If
            Case Else
                GetVersion = "Failed"
        End Select
    End With
End Function

