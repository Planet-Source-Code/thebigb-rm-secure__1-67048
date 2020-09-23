VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Locked"
   ClientHeight    =   2460
   ClientLeft      =   285
   ClientTop       =   390
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLock.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timCD 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6180
      Tag             =   "0"
      Top             =   930
   End
   Begin VB.CommandButton cmdPass 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   4395
      TabIndex        =   3
      Top             =   1095
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "l"
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.ListBox lstHidden 
      Height          =   510
      Left            =   5775
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Timer timHide 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6180
      Top             =   510
   End
   Begin VB.Label lblInfo1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "When the key is inserted tap Enter"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1560
      TabIndex        =   0
      Top             =   750
      Width           =   3285
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IEnumWindowsSink
Dim frmhWnd As Long
Dim FirstTime As Boolean

'// Check if password is correct if used
Private Sub cmdPass_Click()

    If txtPass.Text = Password Then
        timHide.Enabled = False
        ShowAllWnd lstHidden
        Unload Me
    End If

End Sub

'// Initialize the locking
Private Sub Form_Load()

    FirstTime = True
    EnumerateWindows Me
    frmhWnd = Me.hwnd
    DoEvents
    lblInfo1.Caption = "Locking computer..."
    timHide.Enabled = True
    timCD.Enabled = True

End Sub

'// Part of enumeration thing
Private Property Get IEnumWindowsSink_Identifier() As Long

    IEnumWindowsSink_Identifier = Me.hwnd
    
End Property

'// Enumerate all windows and show which ones are visible
'// Pulled from VbAccelerator
'//   http://www.vbaccelerator.com/home/vb/code/Libraries/Windows/Enumerating_Windows/article.asp
Private Sub IEnumWindowsSink_EnumWindow(ByVal hwnd As Long, bStop As Boolean)
On Error Resume Next

    Dim IsFound As Boolean
    Dim i As Long
    
    If hwnd <> frmhWnd Then
        If IsWindowVisible(hwnd) = 1 Then
            For i = 0 To lstHidden.ListCount - 1
                    IsFound = False
                    If lstHidden.List(i) = hwnd Then
                        IsFound = True
                        Exit For
                    End If
            Next i

            If IsFound = False Then
                WindowControl hwnd, HideWnd
                lstHidden.AddItem hwnd
                Exit Sub
            End If
        End If
    Else
        SetTopMost frmhWnd
    End If

End Sub

'// Countdown to prepare the form
Private Sub timCD_Timer()

    If Val(timCD.Tag) < 3 Then
        timCD.Tag = Val(timCD.Tag) + 1
    End If
    
    If Val(timCD.Tag) = 3 Then
        lblInfo1.Caption = "When the key is inserted tap Enter"
        FirstTime = False
        timCD.Enabled = False
    End If

End Sub

'// Make sure every window except itself gets hidden.
'// To restore everything properly, we put hidden objects in a list
Private Sub timHide_Timer()

    EnumerateWindows Me
    If GetAsyncKeyState(13) <> 0 Then
        If CheckDrives = True Then
            If PassUse = True Then
                lblInfo1.Visible = False
                txtPass.Visible = True
                cmdPass.Visible = True
            Else
                timHide.Enabled = False
                ShowAllWnd lstHidden
                Unload Me
            End If
        End If
    End If

End Sub

'// Function to check wether a valid drive is inserted
Private Function CheckDrives() As Boolean

    Dim fso
    Dim Drive
    Dim DrvChkD(2) As Boolean
    Dim DrvChk As Boolean

    DrvChkD(1) = False
    DrvChkD(2) = False
    DrvChk = False

    '// Prepare the FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    DoEvents
    
    If FirstTime = False Then
        lblInfo1.Caption = "Searching keys..."
    End If
    
    '// This part retrieves the serial numbers from the removable media
    For Each Drive In fso.Drives
        If Drive.IsReady = True Then
            If DualDrive = True Then
                If Drive.SerialNumber = SelDrvSerD(1) Then
                    DrvChkD(1) = True
                End If
                
                If Drive.SerialNumber = SelDrvSerD(2) Then
                    DrvChkD(2) = True
                End If
            Else
                If Drive.SerialNumber = SelDrvSer Then
                    DrvChk = True
                End If
            End If
        End If
    Next
    
    '// Here it will compare the serialnumbers
    If FirstTime = False Then
        If DualDrive = True Then
            If DrvChkD(1) = False Or DrvChkD(2) = False Then
                lblInfo1.Caption = "When the key is inserted tap Enter"
                MsgBox ("Key Missing"), vbExclamation, "Key Security"
                CheckDrives = False
                Exit Function
            End If
            CheckDrives = True
        Else
            If DrvChk = False Then
                lblInfo1.Caption = "When the key is inserted tap Enter"
                MsgBox ("Key Missing"), vbExclamation, "Key Security"
                CheckDrives = False
            Else
                CheckDrives = True
            End If
        End If
    Else
        FirstTime = False
    End If
    
End Function

'// Show all windows that were hidden
Private Function ShowAllWnd(WndList As Object)

    Dim i As Long
    
    For i = 0 To WndList.ListCount - 1
        WindowControl WndList.List(i), ShowWnd
    Next i

End Function
