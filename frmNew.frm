VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00E29B81&
   Caption         =   "New Account"
   ClientHeight    =   6165
   ClientLeft      =   3060
   ClientTop       =   1350
   ClientWidth     =   6645
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6645
   Begin VB.Frame Frame2 
      BackColor       =   &H00E29B81&
      Height          =   1935
      Left            =   2760
      TabIndex        =   31
      Top             =   4080
      Width           =   3735
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00E29B81&
         Caption         =   "Previous"
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00E29B81&
         Caption         =   "Next"
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E29B81&
         Caption         =   "Cancel"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00E29B81&
         Caption         =   "Create Account"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Rage Italic"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Rage Italic"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   495
         Left            =   1800
         TabIndex        =   39
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inderpal"
         BeginProperty Font 
            Name            =   "Rage Italic"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   555
         Left            =   1800
         TabIndex        =   38
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E29B81&
      Height          =   6135
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtRelation 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   12
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox txtComment 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtSite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtFirst 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtZip 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.ComboBox ComboSex 
         Height          =   315
         Left            =   4440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "0"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Tag             =   "0"
         Top             =   2400
         Width           =   5415
      End
      Begin VB.TextBox txtLast 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Tag             =   "0"
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E29B81&
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   6315
         TabIndex        =   19
         Top             =   240
         Width           =   6375
         Begin VB.TextBox txtAccount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   0
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3600
            TabIndex        =   23
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   22
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblAccount 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address Id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   36
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   35
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Homepage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "ZipCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmNew"
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

Public db As Database
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
Dim AccCnt As String
Dim AccNum As Integer
    If cmdCreate.Caption = "Create Address" Then
    cmdCreate.Enabled = True
    If txtLast.Text = "" Then
        MsgBox "Please Enter Name"
        txtLast.SetFocus
        Exit Sub
    End If
    If txtFirst.Text = "" Then
        MsgBox "Please Enter Name"
        txtFirst.SetFocus
        Exit Sub
    End If
    If InStr(1, txtMail, "@") = 0 Then
        MsgBox "The Senders email address must contain an @ character"
        txtMail.SetFocus
        Exit Sub
    End If
    If txtAge.Text = "" Then
        MsgBox "Please Enter Age"
        txtAge.SetFocus
        Exit Sub
    End If
    If ComboSex.Text = "" Then
        MsgBox "Please Enter Sex"
        ComboSex.SetFocus
        Exit Sub
    End If
    
    If txtAddress.Text = "" Then
        MsgBox "Please Enter Address"
        txtAddress.SetFocus
        Exit Sub
    End If
    If txtCity.Text = "" Then
        MsgBox "Please Enter City"
        txtCity.SetFocus
        Exit Sub
    End If
    If txtState.Text = "" Then
        MsgBox "Please Enter City"
        txtState.SetFocus
        Exit Sub
    End If
    If txtZip.Text = "" Then
        MsgBox "Please Enter ZipCode"
        txtZip.SetFocus
        Exit Sub
    End If
        With rst1
            .AddNew
            !AddressId = lblAccount.Caption
            !Date = lblDate.Caption
            !LastName = txtLast.Text
            !FirstName = txtFirst.Text
            !EMail = txtMail.Text
            If txtSite.Text = "" Then
            !HomePage = ""
            Else
                !HomePage = txtSite.Text
            End If
            !Age = txtAge.Text
            !Sex = ComboSex.List(ComboSex.ListIndex)
            !Address1 = txtAddress.Text
            !city = txtCity.Text
            !State = txtState.Text
            !ZipCode = txtZip.Text
            If txtPhone = "" Then
                !Phone = ""
            Else
                !Phone = txtPhone.Text
            End If
            If txtRelation.Text = "" Then
                !Relation = ""
            Else
                !Relation = txtRelation.Text
            End If
            If txtComment.Text = "" Then
                !Comments = ""
            Else
                !Comments = txtComment.Text
            End If
            .Update
        End With
        AccCnt = lblAccount.Caption
        AccNum = Val(AccCnt)
        AccNum = AccNum + 1
        Open "Accno.cnt" For Output As #1
            Write #1, AccNum
        Close #1
    End If
    If cmdCreate.Caption = "Modify Address" Then
        With rst
        On Error Resume Next
            !AddressId = txtAccount
            !Date = lblDate.Caption
            !LastName = txtLast.Text
            !FirstName = txtFirst.Text
            !EMail = txtMail.Text
            If txtSite.Text = "" Then
            !HomePage = ""
            Else
                !HomePage = txtSite.Text
            End If
            !Age = txtAge.Text
            !Sex = ComboSex.Text
            !Address1 = txtAddress.Text
            !city = txtCity.Text
            !State = txtState.Text
            !ZipCode = txtZip.Text
            If txtPhone = "" Then
                !Phone = ""
            Else
                !Phone = txtPhone.Text
            End If
            If txtRelation.Text = "" Then
                !Relation = ""
            Else
                !Relation = txtRelation.Text
            End If
            If txtComment.Text = "" Then
                !Comments = ""
            Else
                !Comments = txtComment.Text
            End If
            .Update
        End With
    End If
    If cmdCreate.Caption = "Delete Address" Then
        rst.Delete
    End If
    If cmdCreate.Caption = "Main Menu" Then
    End If
    Unload Me
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    With rst1
    .MoveNext
    If rst1.EOF Then
        MsgBox "End of file.", vbOKOnly, "  End!"
        .MoveLast
        Exit Sub
    Else
            txtAccount = !AddressId
            lblDate = !Date
            txtLast = !LastName
            txtFirst = !FirstName
            txtMail = !EMail
            txtSite = !HomePage
            txtAge = !Age
            ComboSex = !Sex
            txtAddress = !Address1
            txtCity = !city
            txtState = !State
            txtZip = !ZipCode
            txtPhone = !Phone
            txtRelation = !Relation
            txtComment = !Comments
    End If
End With
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    With rst1
    .MovePrevious
    If rst1.BOF Then
        MsgBox "Front of file.", vbOKOnly, "  End!"
        .MoveFirst
        Exit Sub
    Else
            txtAccount = !AddressId
            lblDate = !Date
            txtLast = !LastName
            txtFirst = !FirstName
            txtMail = !EMail
            txtSite = !HomePage
            txtAge = !Age
            ComboSex = !Sex
            txtAddress = !Address1
            txtCity = !city
            txtState = !State
            txtZip = !ZipCode
            txtPhone = !Phone
            txtRelation = !Relation
            txtComment = !Comments
    End If
End With
End Sub

Private Sub Form_Load()
    Dim AccCnt As String * 5
    Open "AccNo.cnt" For Input As #1
        Line Input #1, AccCnt
    Close #1
    
    lblAccount.Caption = AccCnt
    lblDate.Caption = Date
    Set db = OpenDatabase(App.Path & "\Address.mdb")
    With db
        Set rst1 = .OpenRecordset("Address")
    End With
    ComboSex.AddItem "Male"
    ComboSex.AddItem "Female"
    
    Set cnn1 = New ADODB.Connection
    cnn1.Provider = "Microsoft.Jet.OLEDB.3.51"
    cnn1.Open App.Path & "\Address.mdb", "Admin"
    If Flag = True Then
        With rst1
            txtAccount = !AddressId
            lblDate = !Date
            txtLast = !LastName
            txtFirst = !FirstName
            txtMail = !EMail
            txtSite = !HomePage
            txtAge = !Age
            ComboSex = !Sex
            txtAddress = !Address1
            txtCity = !city
            txtState = !State
            txtZip = !ZipCode
            txtPhone = !Phone
            txtRelation = !Relation
            txtComment = !Comments
        End With
    End If
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
    Dim Val1 As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
       Dim strsql As String
        If (rst.State = adStateOpen) Then
            rst.Close
            rst.ActiveConnection = Nothing
        End If
        strsql = "SELECT * FROM Address WHERE AddressId='" & txtAccount.Text & "'"
        With rst
            rst.ActiveConnection = cnn1
            rst.CursorType = adOpenKeyset
            rst.LockType = adLockOptimistic
            rst.Open strsql
            Val1 = Val(!AddressId)
            If Val1 = Val(txtAccount.Text) Then
                cmdCreate.Enabled = True
                lblDate = !Date
                txtLast = !LastName
                txtFirst = !FirstName
                txtMail = !EMail
                txtSite = !HomePage
                txtAge = !Age
                ComboSex = !Sex
                txtAddress = !Address1
                txtCity = !city
                txtState = !State
                txtZip = !ZipCode
                txtPhone = !Phone
                txtRelation = !Relation
                txtComment = !Comments
            Else
                Call Clear
                cmdCreate.Enabled = False
                txtAccount.Text = ""
                MsgBox "Address Deleted or Not Created", vbInformation, "Address Not Found"
                txtAccount.SetFocus
            End If
        End With
    End If
    If cmdCreate.Caption = "Modify Address" Then
        Call Disable
    End If
End Sub
Private Sub Disable()
    frmNew.txtLast.Enabled = False
    frmNew.txtFirst.Enabled = False
    frmNew.txtAge.Enabled = False
    frmNew.ComboSex.Enabled = False
End Sub

Private Sub Clear()
    txtLast.Text = ""
    txtFirst.Text = ""
    txtMail = ""
    txtSite = ""
    txtAge = ""
    ComboSex.List(ComboSex.ListIndex) = ""
    txtAddress = ""
    txtCity = ""
    txtState = ""
    txtZip = ""
    txtPhone = ""
    txtRelation = ""
    txtComment = ""
End Sub
