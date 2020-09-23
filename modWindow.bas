Attribute VB_Name = "modWindow"
'============================================================================================================
'   Pulled from VbAccelerator
'   http://www.vbaccelerator.com/home/vb/code/Libraries/Windows/Enumerating_Windows/article.asp
'============================================================================================================

Option Explicit

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, _
    ByVal cch As Long) As Long
    
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Const WM_COMMAND = &H111

Private m_cSink As IEnumWindowsSink

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

    Dim bStop As Boolean
    
    bStop = False
    m_cSink.EnumWindow hwnd, bStop
    If (bStop) Then
        EnumWindowsProc = 0
    Else
        EnumWindowsProc = 1
    End If
    
End Function

Public Function EnumerateWindows(ByRef cSink As IEnumWindowsSink) As Boolean

    If Not (m_cSink Is Nothing) Then Exit Function
    Set m_cSink = cSink
    EnumWindows AddressOf EnumWindowsProc, cSink.Identifier
    Set m_cSink = Nothing

End Function

Public Function WindowTitle(ByVal lHwnd As Long) As String

    Dim lLen As Long
    Dim sBuf As String

    ' Get the Window Title:
    lLen = GetWindowTextLength(lHwnd)
    If (lLen > 0) Then
        sBuf = String$(lLen + 1, 0)
        lLen = GetWindowText(lHwnd, sBuf, lLen + 1)
        WindowTitle = Left$(sBuf, lLen)
    End If
    
End Function

Public Function ClassName(ByVal lHwnd As Long) As String
    
    Dim lLen As Long
    Dim sBuf As String
    
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
    
End Function

Public Sub ActivateWindow(ByVal lHwnd As Long)

    SetForegroundWindow lHwnd
    
End Sub
