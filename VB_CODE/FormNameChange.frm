VERSION 5.00
Begin VB.Form FormNameChange 
   Caption         =   "Change Secure Folder Name"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4350
   Icon            =   "FormNameChange.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2130
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Secure Folder Name :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "FormNameChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim user As String, pass As String, l As Integer, prev As String
    Me.Hide
    l = 0
    On Error GoTo error_handle_newname
    user = decrypt(read_file(App.Path + "/Folder_lock_data/Username.txt"))
    l = 1
    pass = decrypt(read_file(App.Path + "/Folder_lock_data/Password.txt"))
    MsgBox "Secure Folder Was Already Created . It Is Not Possible To Change The Secure Folder Name ...!!!", vbCritical, "Error In File Creation"
    Text1.Text = ""
If l = 0 Then
error_handle_newname:
    prev = decrypt(read_file(App.Path + "/Folder_lock_data/dirname.txt"))
    x = write_file(App.Path + "/Folder_lock_data/dirname.txt", encrypt(Text1.Text))
    MsgBox "Secure Folder Name Has Been Changed From " + prev + " To " + Text1.Text, vbOKOnly
    Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
    Me.Hide
    MsgBox "Secure Folder Name Change Is Interrupted", vbOKOnly
End Sub

