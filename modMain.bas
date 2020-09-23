Attribute VB_Name = "modMain"
Option Explicit
 
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, _
    ByVal X&, ByVal Y&, ByVal cX&, ByVal cY&, ByVal wFlags&)
    
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Password As String
Public SelDrv As String
Public SelDrvSer As String
Public SelDrvD(2) As String
Public SelDrvSerD(2) As String
Public DualDrive As Boolean
Public PassUse As Boolean

Const SW_HIDE = 0
Const SW_SHOW = 5

Public Const VK_SHIFT As Long = &H10
Public Const VK_CONTROL As Long = &H11
Public Const VK_MENU As Long = &H12
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
   
Public Const HWND_TOPMOST& = -1
Public Const SWP_NOMOVE& = &H2
Public Const SWP_NOSIZE& = &H1
Public Const SWP_NOACTIVATE& = &H10

Enum GS
    GetS = 0
    SetS = 1
End Enum

Enum wFunction
    ShowWnd = 0
    HideWnd = 1
End Enum

Enum ED
    Enable = 0
    Disable = 1
End Enum

'// Sets the window on top
Public Sub SetTopMost(ByVal hwnd&)

    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)
    
End Sub

'// Show or hide a window
Public Sub WindowControl(hwnd As Long, WindowFunction As wFunction)

    Select Case WindowFunction
        Case ShowWnd
            ShowWindow hwnd, SW_SHOW
        Case HideWnd
            ShowWindow hwnd, SW_HIDE
    End Select

End Sub

'// Get or set the settings
'// OK, this is a weak way to save the password. If you intend to
'// use this as real security, you may want to add a little encryption.
Public Function GSSettings(GetOrSet As GS, DoActions As Boolean)
    
    If GetOrSet = GetS Then
        If GetSetting("KeySecure", "a", "pwuse") = "1" Then
            PassUse = True
        Else
            PassUse = False
        End If
        
        Password = GetSetting("KeySecure", "a", "pw")
        
        If GetSetting("KeySecure", "a", "dd") = "1" Then
            DualDrive = True
        Else
            DualDrive = False
        End If
        
        SelDrv = GetSetting("KeySecure", "a", "sd")
        SelDrvSer = GetSetting("KeySecure", "a", "sds")
        SelDrvD(1) = GetSetting("KeySecure", "a", "sd1")
        SelDrvD(2) = GetSetting("KeySecure", "a", "sd2")
        SelDrvSerD(1) = GetSetting("KeySecure", "a", "sd1s")
        SelDrvSerD(2) = GetSetting("KeySecure", "a", "sd2s")

        If DoActions = True Then
            If DualDrive = True Then
                frmSettings.chkSecondary.Value = Checked
            Else
                frmSettings.chkSecondary.Value = Unchecked
            End If
            
            If PassUse = True Then
                frmSettings.chkPass.Value = Checked
            Else
                frmSettings.chkPass.Value = Unchecked
            End If
        End If
    Else
        If PassUse = True Then
            SaveSetting "KeySecure", "a", "pwuse", "1"
        Else
            SaveSetting "KeySecure", "a", "pwuse", "0"
        End If
        
        If DualDrive = True Then
            SaveSetting "KeySecure", "a", "dd", "1"
        Else
            SaveSetting "KeySecure", "a", "dd", "0"
        End If
                
        SaveSetting "KeySecure", "a", "pw", Password
        SaveSetting "KeySecure", "a", "sd", SelDrv
        SaveSetting "KeySecure", "a", "sds", SelDrvSer
        SaveSetting "KeySecure", "a", "sd1", SelDrvD(1)
        SaveSetting "KeySecure", "a", "sd2", SelDrvD(2)
        SaveSetting "KeySecure", "a", "sds1", SelDrvSerD(1)
        SaveSetting "KeySecure", "a", "sds2", SelDrvSerD(2)
    End If
            
End Function

'// ** I don't use this in my code simply because it doesn't work... **
'// But in theory this would block combinations like ctrl+alt+del.
'// The idea is, we hold down the shift key. Try to hit ctrl+alt+del again, and
'// nothing will happen.
Public Function BlockLoKeys(State As ED)

    If State = Enable Then
        If GetAsyncKeyState(VK_CONTROL) <> 0 Then
            keybd_event VK_SHIFT, 0, 0, 0
            DoEvents
        End If
        
        If GetAsyncKeyState(VK_MENU) <> 0 Then
            keybd_event VK_CONTROL, 0, 0, 0
            keybd_event VK_SHIFT, 0, 0, 0
            DoEvents
        End If

    End If
    
    If State = Disable Then
        keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
        keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    End If
    
    'Example:
    '========
    'keybd_event VK_H, 0, 0, 0                 ' press H
    'keybd_event VK_H, 0, KEYEVENTF_KEYUP, 0   ' release H

End Function
