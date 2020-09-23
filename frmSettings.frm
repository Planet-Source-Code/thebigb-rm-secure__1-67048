VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5100
   ClientLeft      =   285
   ClientTop       =   390
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList im 
      Left            =   1380
      Top             =   4695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3452
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":37A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3AF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":419A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer timResize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   900
      Tag             =   "Grow"
      Top             =   4695
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   4725
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2895
      TabIndex        =   8
      Top             =   4725
      Width           =   780
   End
   Begin VB.Frame frPass 
      Caption         =   "Password"
      ForeColor       =   &H00C00000&
      Height          =   885
      Left            =   60
      TabIndex        =   5
      Top             =   3765
      Width           =   3615
      Begin VB.CheckBox chkPass 
         Caption         =   "Use Password"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   375
         Width           =   2160
      End
      Begin VB.CommandButton cmdPass 
         Caption         =   "Set/Change"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2340
         TabIndex        =   6
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame frDrive 
      Caption         =   "Drive"
      ForeColor       =   &H00C00000&
      Height          =   3660
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3615
      Begin VB.CheckBox chkSecondary 
         Caption         =   "Use secondary drive"
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   3300
         Width           =   2445
      End
      Begin VB.CommandButton cmdSelect2 
         Caption         =   "Select"
         Height          =   300
         Left            =   2550
         TabIndex        =   3
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh List"
         Height          =   300
         Left            =   150
         TabIndex        =   2
         Top             =   2520
         Width           =   1050
      End
      Begin MSComctlLib.TreeView tvDrives 
         Height          =   2190
         Left            =   150
         TabIndex        =   1
         Top             =   315
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3863
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   647
         Style           =   7
         ImageList       =   "im"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdSelect1 
         Caption         =   "Select (1)"
         Height          =   300
         Left            =   1710
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSelDrv 
         Caption         =   "Selected Drive: - "
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   2970
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   2040
      TabIndex        =   9
      Top             =   4725
      Width           =   870
   End
   Begin VB.Label lblHelp 
      Caption         =   $"frmSettings.frx":44EC
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   3795
      TabIndex        =   11
      Top             =   300
      Width           =   2445
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum VF
    Visible = 0
    Enabled = 1
End Enum

'// If the checkbox is checked enable/disable the 'edit' button
Private Sub chkPass_Click()

    CheckChk chkPass, cmdPass, Enabled
    
    If chkPass.Value = Checked Then
        PassUse = True
    Else
        PassUse = False
    End If

End Sub

'// Same as above
Private Sub chkPass_KeyPress(KeyAscii As Integer)

    chkPass_Click

End Sub

'// When you choose for a secondary drive, a second button will appear
Private Sub chkSecondary_Click()

    CheckChk chkSecondary, cmdSelect1, Visible
    
    If chkSecondary.Value = Checked Then
        cmdSelect2.Caption = "Select (2)"
        lblSelDrv.Caption = "Selected Drives: -"
        DualDrive = True
    Else
        cmdSelect2.Caption = "Select"
        lblSelDrv.Caption = "Selected Drive: -"
        DualDrive = False
    End If

End Sub

'// Same as above
Private Sub chkSecondary_KeyPress(KeyAscii As Integer)

    chkSecondary_Click

End Sub

'// Unload the form
Private Sub cmdCancel_Click()

    Unload Me

End Sub

'// Resize form to show/hide label
Private Sub cmdHelp_Click()

    timResize.Enabled = True

End Sub

'// Save settings and lock the computer
Private Sub cmdOK_Click()

    GSSettings SetS, False
    DoEvents
    frmLock.Show
    Unload Me

End Sub

'// Show password edit dialog
Private Sub cmdPass_Click()

    frmPassword.Show

End Sub

'// Refresh drives list
Private Sub cmdRefresh_Click()

    LoadDrives tvDrives

End Sub

'// If this button is clicked we deal with dualmedia.
'// Here it will set the first drive
Private Sub cmdSelect1_Click()
On Error Resume Next

    SelDrvD(1) = tvDrives.SelectedItem.Text

    If Left(tvDrives.SelectedItem.Key, 3) = "SER" Then
        SelDrvSerD(1) = Split(tvDrives.SelectedItem.Key, "SER")(1)
    Else
        SelDrvSerD(1) = 0
    End If
    
    If SelDrvD(2) <> "" Then
        lblSelDrv.Caption = "Selected Drives: " & Chr(34) & SelDrvD(1) & Chr(34) & " && " & _
            Chr(34) & SelDrvD(2) & Chr(34)
    Else
        lblSelDrv.Caption = "Selected Drives: " & Chr(34) & SelDrvD(1) & Chr(34)
    End If

End Sub

'// Set second or only drive
Private Sub cmdSelect2_Click()
On Error Resume Next

    If chkSecondary.Value = Checked Then
        SelDrvD(2) = tvDrives.SelectedItem.Text
        If Left(tvDrives.SelectedItem.Key, 3) = "SER" Then
            SelDrvSerD(2) = Split(tvDrives.SelectedItem.Key, "SER")(1)
        Else
            SelDrvSerD(2) = 0
        End If
        
        lblSelDrv.Caption = "Selected Drives: " & Chr(34) & SelDrvD(1) & Chr(34) & " && " & _
            Chr(34) & SelDrvD(2) & Chr(34)

    Else
        SelDrv = tvDrives.SelectedItem.Text
        If Left(tvDrives.SelectedItem.Key, 3) = "SER" Then
            SelDrvSer = Split(tvDrives.SelectedItem.Key, "SER")(1)
        Else
            SelDrvSer = 0
        End If
        lblSelDrv.Caption = "Selected Drive: " & Chr(34) & SelDrv & Chr(34)
    End If

End Sub

'// Load drives list and get settings
Private Sub Form_Load()

    LoadDrives tvDrives
    GSSettings GetS, True

End Sub

'// Resize form for help
Private Sub timResize_Timer()

    If timResize.Tag = "Grow" Then
        If Me.Width < 6510 Then
            Me.Caption = ""
            Me.Width = Me.Width + 160
        End If
        If Me.Width = 6510 Or Me.Width > 6510 Then
            Me.Width = 6510
            Me.Caption = "Settings"
            timResize.Tag = "Shrink"
            timResize.Enabled = False
        End If
    Else
        If Me.Width > 3825 Then
            Me.Caption = ""
            Me.Width = Me.Width - 60
        End If
        If Me.Width = 3825 Or Me.Width < 3825 Then
            Me.Width = 3825
            Me.Caption = "Settings"
            timResize.Tag = "Grow"
            timResize.Enabled = False
        End If
    End If
    
End Sub

'// Get the drives
Private Function LoadDrives(Treeview As Object)

    Dim fso
    Dim Drive
    
    Treeview.Nodes.Clear
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    DoEvents
    
    For Each Drive In fso.Drives
        If Drive.IsReady = True Then
            '0 = Unknown
            '1 = Removable
            '2 = Fixed
            '3 = Remote
            '4 = Cd-Rom
            '5 = Ram Disk
            DoEvents
            Select Case Drive.DriveType
                Case 0
                    If Drive.SerialNumber <> 0 Then
                        Treeview.Nodes.Add , , "SER" & Drive.SerialNumber, Drive.DriveLetter & ":\", 1
                    Else
                        Treeview.Nodes.Add , , Drive.DriveLetter, Drive.DriveLetter & ":\", 1
                    End If
                Case 1
                    If Drive.SerialNumber <> 0 Then
                        '// You can change the maximum size of removable media to own
                        '// will. This is here to seperate external and internal disks
                        '// but as the years fly by there is no much difference anymore...
                        If Drive.TotalSize < 3000000 Then
                            Treeview.Nodes.Add , , "SER" & Drive.SerialNumber, Drive.DriveLetter & _
                                ":\", 4
                        Else
                            Treeview.Nodes.Add , , "SER" & Drive.SerialNumber, Drive.DriveLetter & _
                                ":\", 1
                        End If
                    Else
                        If Drive.TotalSize < 3000000 Then
                            Treeview.Nodes.Add , , Drive.DriveLetter, Drive.DriveLetter & _
                                ":\", 4
                        Else
                            Treeview.Nodes.Add , , Drive.DriveLetter, Drive.DriveLetter & _
                                ":\", 1
                        End If
                    End If
                Case 3
                    If Drive.SerialNumber <> 0 Then
                        Treeview.Nodes.Add , , "SER" & Drive.SerialNumber, Drive.DriveLetter & ":\", 5
                    Else
                        Treeview.Nodes.Add , , Drive.DriveLetter, Drive.DriveLetter & ":\", 5
                    End If
                Case 4
                    If Drive.SerialNumber <> 0 Then
                        Treeview.Nodes.Add , , "SER" & Drive.SerialNumber, Drive.DriveLetter & ":\", 2
                    Else
                        Treeview.Nodes.Add , , Drive.DriveLetter, Drive.DriveLetter & ":\", 2
                    End If
                Case 5
                    If Drive.SerialNumber <> 0 Then
                        Treeview.Nodes.Add , , "SER" & Drive.SerialNumber, Drive.DriveLetter & ":\", 1
                    Else
                        Treeview.Nodes.Add , , Drive.DriveLetter, Drive.DriveLetter & ":\", 1
                    End If
            End Select
        End If
    Next
    
End Function

Private Function CheckChk(ChkBox As Object, CmdButton As Object, VisibleOrEnabled As VF)
On Error Resume Next

    If ChkBox.Value = Checked Then
        If VisibleOrEnabled = Enabled Then
            CmdButton.Enabled = True
        Else
            CmdButton.Visible = True
        End If
    Else
        If VisibleOrEnabled = Enabled Then
            CmdButton.Enabled = False
        Else
            CmdButton.Visible = False
        End If
    End If
    
End Function

