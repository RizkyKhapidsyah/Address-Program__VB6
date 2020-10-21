VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E29B81&
   Caption         =   "Address Book"
   ClientHeight    =   4035
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   3000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3000
   Begin VB.Frame Frame1 
      BackColor       =   &H00E29B81&
      Caption         =   "Please Select An Option"
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H00E29B81&
         Caption         =   "About Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00E29B81&
         Caption         =   "View Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E29B81&
         Caption         =   "Exiting Program . . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00E29B81&
         Caption         =   "Find Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E29B81&
         Caption         =   "Delete Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00E29B81&
         Caption         =   "Modify Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00E29B81&
         Caption         =   "Create New Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        '########################################'
        '   Programmed By Inderpal Singh         '
        '   Email: inderpal0@hotmail.com         '
        '   Date: March 23, 2002                 '
        '   Homepage: http://connect.to/lanserver'
        '########################################'

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdDelete_Click()
    Flag = False
    Me.Hide
    frmNew.Caption = "Delete Address"
    frmNew.cmdCreate.Caption = "Delete Address"
    frmNew.txtAccount.Visible = True
    frmNew.lblAccount.Visible = False
    Call LabelEnable
    frmNew.Show vbModal
    Me.Show
End Sub

Private Sub cmdDisplay_Click()
    Flag = True
    Me.Hide
    frmNew.cmdCreate.Caption = "Main Menu"
    frmNew.Caption = "Address Book"
    frmNew.txtAccount.Visible = True
    frmNew.txtAccount.Enabled = False
    frmNew.lblAccount.Visible = False
    Call CmdEnable
    frmNew.Show vbModal
    Me.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdModify_Click()
    Flag = False
    Me.Hide
    frmNew.Caption = "Modify Address"
    frmNew.cmdCreate.Caption = "Modify Address"
    frmNew.txtAccount.Visible = True
    frmNew.lblAccount.Visible = False
    Call LabelEnable
    frmNew.Show vbModal
    Me.Show
End Sub

Private Sub cmdNew_Click()
    Flag = False
    Me.Hide
    frmNew.Caption = "Create Account"
    frmNew.cmdCreate.Caption = "Create Address"
    frmNew.txtAccount.Visible = False
    frmNew.lblAccount.Visible = True
    Call LabelEnable
    frmNew.Show vbModal
    Me.Show
End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdNew.BackColor = &HFF&
    cmdModify.BackColor = &HE29B81
    cmdDelete.BackColor = &HE29B81
    cmdView.BackColor = &HE29B81
    cmdDisplay.BackColor = &HE29B81
    cmdExit.BackColor = &HE29B81
    cmdAbout.BackColor = &HE29B81
End Sub
Private Sub cmdModify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdModify.BackColor = &HFF&
    cmdNew.BackColor = &HE29B81
    cmdDelete.BackColor = &HE29B81
    cmdView.BackColor = &HE29B81
    cmdDisplay.BackColor = &HE29B81
    cmdExit.BackColor = &HE29B81
    cmdAbout.BackColor = &HE29B81
End Sub
Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdDelete.BackColor = &HFF&
    cmdModify.BackColor = &HE29B81
    cmdNew.BackColor = &HE29B81
    cmdView.BackColor = &HE29B81
    cmdDisplay.BackColor = &HE29B81
    cmdExit.BackColor = &HE29B81
    cmdAbout.BackColor = &HE29B81
End Sub
Private Sub cmdView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdView.BackColor = &HFF&
    cmdModify.BackColor = &HE29B81
    cmdNew.BackColor = &HE29B81
    cmdDelete.BackColor = &HE29B81
    cmdDisplay.BackColor = &HE29B81
    cmdExit.BackColor = &HE29B81
    cmdAbout.BackColor = &HE29B81
End Sub
Private Sub cmdDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdDisplay.BackColor = &HFF&
    cmdModify.BackColor = &HE29B81
    cmdNew.BackColor = &HE29B81
    cmdDelete.BackColor = &HE29B81
    cmdView.BackColor = &HE29B81
    cmdExit.BackColor = &HE29B81
    cmdAbout.BackColor = &HE29B81
End Sub
Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdExit.BackColor = &HE29B81
    cmdModify.BackColor = &HE29B81
    cmdNew.BackColor = &HE29B81
    cmdDelete.BackColor = &HE29B81
    cmdView.BackColor = &HE29B81
    cmdDisplay.BackColor = &HE29B81
    cmdAbout.BackColor = &HFF&
End Sub
Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdExit.BackColor = &HFF&
    cmdModify.BackColor = &HE29B81
    cmdNew.BackColor = &HE29B81
    cmdDelete.BackColor = &HE29B81
    cmdView.BackColor = &HE29B81
    cmdDisplay.BackColor = &HE29B81
    cmdAbout.BackColor = &HE29B81
End Sub
Private Sub cmdView_Click()
    Flag = False
    Me.Hide
    frmNew.Caption = "View Address"
    frmNew.cmdCreate.Caption = "Main Menu"
    Call LabelEnable
    frmNew.Show vbModal
    Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub LabelEnable()
    frmNew.Label16.Visible = True
    frmNew.Label17.Visible = True
    frmNew.Label18.Visible = True
    frmNew.cmdNext.Visible = False
    frmNew.cmdPrevious.Visible = False
End Sub

Private Sub CmdEnable()
    frmNew.Label16.Visible = False
    frmNew.Label17.Visible = False
    frmNew.Label18.Visible = False
    frmNew.cmdNext.Visible = True
    frmNew.cmdPrevious.Visible = True
End Sub
