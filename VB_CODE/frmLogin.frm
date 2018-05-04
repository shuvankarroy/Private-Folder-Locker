VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Username & Password"
   ClientHeight    =   1845
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1090.087
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1320
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Current Username and Password :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   960
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer

Private Sub cmdCancel_Click()
    frmLogin.Visible = False
End Sub

Private Sub cmdOK_Click()
    frmLogin.Visible = False
    Dim user As String, pass As String
    On Error GoTo error_usr_change
    
    user = frmLogin.txtUserName
    pass = frmLogin.txtPassword
    txtUserName.Text = ""
    txtPassword.Text = ""
    If (check(user, pass) = 1 And k = 1) Then
        k = 2
        frmLogin.txtUserName = ""
        frmLogin.txtPassword = ""
        Label1.Caption = "Enter New Username and Password"
        frmLogin.Show
    Else
        If (k = 2) Then
            k = 3
            pass1 = write_file(App.Path + "/Folder_lock_data/UserName.txt", encrypt(user))
            pass2 = write_file(App.Path + "/Folder_lock_data/Password.txt", encrypt(pass))
            MsgBox "Username and Password changed Successfully", vbOKOnly, "Success"
        Else
            k = 3
            frmLogin.txtUserName = ""
            frmLogin.txtPassword = ""
            MsgBox "Wrong Username or Password . Change Username & Password Incomplete . Provide Correct Username And Password And Try Again...!!!", vbCritical, "Incomplete Action"
        End If
    End If
If (k = 1) Then
error_usr_change:
    frmLogin.txtUserName = ""
    frmLogin.txtPassword = ""
    MsgBox "Secure Folder Was Not Created. Create Secure Folder First and Then Try Again .....!!!", vbCritical, "Incomplete Action"
End If
End Sub

Private Sub Form_Load()
    k = 1
End Sub
