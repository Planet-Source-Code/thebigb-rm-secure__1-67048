VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change your password"
   ClientHeight    =   2205
   ClientLeft      =   6690
   ClientTop       =   2535
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassword.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNPass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   315
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1185
      Width           =   3945
   End
   Begin VB.Frame frControls 
      Caption         =   "Key Secure"
      ForeColor       =   &H80000011&
      Height          =   675
      Left            =   -30
      TabIndex        =   7
      Top             =   1575
      Width           =   5115
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   300
         Left            =   4125
         TabIndex        =   2
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   3390
         TabIndex        =   3
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame frLine 
      Height          =   30
      Left            =   -60
      TabIndex        =   6
      Top             =   750
      Width           =   5490
   End
   Begin VB.TextBox txtOPass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   315
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   435
      Width           =   3945
   End
   Begin VB.Label lbl2 
      Caption         =   "New password:"
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   900
      Width           =   4500
   End
   Begin VB.Label lbl1 
      Caption         =   "Old password:"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   4500
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    Unload Me

End Sub

'// Check if old password is correct and change password
Private Sub cmdOK_Click()

    If txtOPass.Text = Password Then
        Password = txtNPass.Text
        Call GSSettings(SetS, False)
        DoEvents
        Unload Me
    Else
        MsgBox ("Invalid Password"), vbExclamation, "Key Security"
        txtOPass.Text = ""
        txtNPass.Text = ""
        txtOPass.SetFocus
    End If

End Sub

