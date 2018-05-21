VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1_MAIN 
   Caption         =   "FOLDER LOCKER"
   ClientHeight    =   4800
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "Form1_MAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "USB MODE"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Readme"
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
      Left            =   4560
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete Secure Folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Create Secure Folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&UNLOCK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remember Me"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&LOCK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Menu f 
      Caption         =   "&File"
      Begin VB.Menu fc 
         Caption         =   "Create Secure Folder"
      End
      Begin VB.Menu fl 
         Caption         =   "Lock"
      End
      Begin VB.Menu fu 
         Caption         =   "Unlock"
      End
      Begin VB.Menu fd 
         Caption         =   "Delete Secure Folder"
      End
      Begin VB.Menu fcn 
         Caption         =   "Change Secure Folder Name"
      End
      Begin VB.Menu fusr 
         Caption         =   "Change Username and Password"
      End
      Begin VB.Menu FSet 
         Caption         =   "Set Secure Folder Path"
      End
      Begin VB.Menu fShow 
         Caption         =   "Show Secure Folder Path"
      End
      Begin VB.Menu fgap 
         Caption         =   "-"
      End
      Begin VB.Menu fe 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu e 
      Caption         =   "&Edit"
      Begin VB.Menu eClear 
         Caption         =   "Clear All Fields"
      End
      Begin VB.Menu eshow 
         Caption         =   "Show Password"
      End
      Begin VB.Menu ehide 
         Caption         =   "Hide Password"
      End
   End
   Begin VB.Menu v 
      Caption         =   "&View"
      Begin VB.Menu vc 
         Caption         =   "Change Background Colour"
      End
   End
   Begin VB.Menu h 
      Caption         =   "&Help"
      Begin VB.Menu ha 
         Caption         =   "About Folder Locker"
      End
   End
End
Attribute VB_Name = "Form1_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dir As String, choice As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Check1_Click()
    'Check Boxes For Remember Password
    Dim user As String, pass As String, i As Integer
    i = 0
    On Error GoTo Handling_Remeber_Error
    user = decrypt(read_file(App.Path + "/Folder_lock_data/Username.txt"))
    pass = decrypt(read_file(App.Path + "/Folder_lock_data/Password.txt"))
    i = 1
    If user = Text1.Text Then
        Text2.Text = pass
    Else
        MsgBox "Wrong Username . Provide Correct Username and Then Try Again...!!!", vbCritical, "Error"
    End If
If i = 0 Then
Handling_Remeber_Error:
    MsgBox "Secure Folder Was Not Created. Create Secure Folder First and Then Try Again .....!!!", vbCritical, "Incomplete Action"
End If
End Sub

Private Sub Check2_Click()
    Dim prev_location As String, cur_location As String, new_location As String, pass1 As Integer, i As Integer, c As Integer
    i = 0
    c = MsgBox("Are You Sure You Want To Turn On USB MODE ? Turn It On Only In Pendrive Or In USB Devices", vbOKCancel, "USB MODE")
    If c = 1 Then
        On Error GoTo Handle_USB_MODE:
        prev_location = decrypt(read_file(App.Path + "/Folder_lock_data/dirlocation.txt"))
        cur_location = App.Path
        new_location = Left(cur_location, 1) + Right(prev_location, Len(prev_location) - 1)
        pass1 = write_file(App.Path + "/Folder_lock_data/dirlocation.txt", encrypt(new_location))
        i = 1
If i = 0 Then
Handle_USB_MODE:
    MsgBox "Please Create The Secure Folder First And Then Turn On USB MODE", vbCritical, "ERROR"
End If
    End If
End Sub

Private Sub Command1_Click()
    'Command Button For Locking Folder
    Dim pass As Integer, i As Integer, py As String, y As Double, pass1 As Integer, pass2 As Integer
    i = 0
    On Error GoTo Handling_Nofile_Error
    pass = check(Text1.Text, Text2.Text)
    If (pass = 1) Then
        pass1 = write_file(Left(dir, 3) + "dirfile.txt", decrypt(read_file(App.Path + "/Folder_lock_data/dirlocation.txt"))) 'for directory location
        pass2 = write_file(Left(dir, 3) + "keyfile.txt", "abrakadabra") 'For sending the key to run python exe
        Shell App.Path + "/Folder_lock_data/Search_and_encrypt.exe", vbHide
        ProgressBar1.Visible = True
        While (isRunningExe("Search_and_encrypt.exe"))
            Sleep (1000)
            If ProgressBar1.Value < 100 Then
                ProgressBar1.Value = ProgressBar1.Value + 10
            End If
            If ProgressBar1.Value = 100 Then
                ProgressBar1.Value = 0
            End If
        Wend
        ProgressBar1.Visible = False
        MsgBox "Secure Folder is Locked", vbOKOnly, "Lock Successful"
        i = 1
    Else
        MsgBox "Wrong Username and Password . Provide Correct Username and Password and Try Again....!!!", vbCritical, "Error"
        i = 1
    End If
If (i = 0) Then
Handling_Nofile_Error:
    MsgBox "Secure Folder Was Not Created. Create Secure Folder First and Then Try Again .....!!!", vbCritical, "Incomplete Action"
End If
End Sub

Private Sub Command2_Click()
    'Command Button For Unlocking Folder
    Dim a As Integer, key As String, i As Integer
    i = 0
    On Error GoTo Handling_Nofile_Error
    a = check(Text1.Text, Text2.Text)
    If (a = 1) Then
        pass1 = write_file(Left(dir, 3) + "dirfile.txt", decrypt(read_file(App.Path + "/Folder_lock_data/dirlocation.txt"))) 'for directory location
        pass2 = write_file(Left(dir, 3) + "keyfile.txt", "abrakadabra") 'for key file
        Shell App.Path + "/Folder_lock_data/Search_and_decrypt.exe", vbHide
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        While (isRunningExe("Search_and_decrypt.exe"))
            Sleep (1000)
            If ProgressBar1.Value < 100 Then
                ProgressBar1.Value = ProgressBar1.Value + 10
            End If
            If ProgressBar1.Value = 100 Then
                ProgressBar1.Value = 0
            End If
        Wend
        ProgressBar1.Visible = False
        MsgBox "Secure Folder is Unlocked", vbOKOnly, "Unlock Successful"
        i = 1
    Else
        MsgBox "Wrong Username and Password . Provide Correct Username and Password and Try Again....!!!", vbCritical, "Error"
        i = 1
    End If
If (i = 0) Then
Handling_Nofile_Error:
    MsgBox "Secure Folder Was Not Created. Create Secure Folder First and Then Try Again .....!!!", vbCritical, "Incomplete Action"
End If
End Sub

Private Sub Command3_Click()
    'Command Button For Creating Secure Folder
    Dim user As String, passwrd As String, i As Integer, pass1 As Integer, pass2 As Integer, pass3 As Integer, pass4 As Integer
    i = 0
    pass1 = 0
    pass2 = 0
    On Error GoTo Create
    user = decrypt(read_file(App.Path + "/Folder_lock_data/Username.txt"))
    i = 1
    passwrd = decrypt(read_file(App.Path + "/Folder_lock_data/Password.txt"))
    MsgBox "Secure Folder Was Already Created....!!!", vbCritical, "Error In File Creation"
If (i = 0) Then
Create:
        If ((Len(Text1.Text) = 0) Or (Len(Text2.Text) = 0)) Then
            MsgBox "Username and Password is Too Short . Provide a Lengthy Username and Password For Better Security .", vbCritical, "Short Username & Password"
        Else
            Call FSet_Click
            If choice = 1 Or choice = 3 Then
                pass1 = write_file(App.Path + "/Folder_lock_data/UserName.txt", encrypt(Text1.Text))
                pass2 = write_file(App.Path + "/Folder_lock_data/Password.txt", encrypt(Text2.Text))
                pass3 = write_file(Left(dir, 3) + "dirfile.txt", decrypt(read_file(App.Path + "/Folder_lock_data/dirlocation.txt"))) 'for Secure directory location
                pass4 = write_file(Left(dir, 3) + "keyfile.txt", "abrakadabra")
                Shell App.Path + "/Folder_lock_data/Create_Folder.exe", vbHide
                DoEvents
                Sleep 1000
                If (pass1 = 1 And pass2 = 1) Then
                    MsgBox "Secure Folder Is Created In Curent Directory And It Is Ready To Use .", vbOKOnly, "Secure Folder Created"
                End If
            Else
                MsgBox "Secure Folder Creation Is Aborted .", vbOKOnly, "Folder Creation Interrupted"
            End If
        End If
End If
End Sub

Private Sub Command4_Click()
    'Command Button For Deleting Secure Folder
    Dim a As Integer, i As Integer, strPath As String
    i = 0
    a = 0
    'On Error GoTo Handling_Deletefile_Error
    a = check(Text1.Text, Text2.Text)
    If (a = 1) Then
        strPath = decrypt(read_file(App.Path + "/Folder_lock_data/dirlocation.txt"))
            inp = MsgBox("Are You Sure That You Want to Permanently Delete Secure Folder?", vbOKCancel, "Permission Required")
            If (inp = 1) Then
                Kill (App.Path + "/Folder_lock_data/Password.txt")
                Kill (App.Path + "/Folder_lock_data/Username.txt")
                RmDir (decrypt(read_file(App.Path + "/Folder_lock_data/dirlocation.txt")))
                Kill (App.Path + "/Folder_lock_data/dirlocation.txt")
                MsgBox "The Secure Folder Is Deleted Permanently . Thank You For Using .", vbOKOnly, "Secure Folder Deleted"
            End If
            i = 1
    Else
        MsgBox "Wrong Username And Password . Secure Folder Deletion Incomplete . Provide Correct Username And Password And Try Again...!!!", vbCritical, "Error In Deletion"
        i = 1
    End If
If (i = 0) Then
Handling_Deletefile_Error:
    MsgBox "Secure Folder Was Not Created or It might be hidden or It may Contain Some Informations . Please Unhide/Unlock The Folder or Create Secure Folder First or Delete All Your Belongings and Then Try Again .....!!!", vbCritical, "Incomplete Action"
End If
End Sub

Private Sub Command5_Click()
    'Command Button To Open readme Notepad
    Shell "notepad " + App.Path + "/Folder_lock_data/Readme.txt", vbNormalFocus
End Sub

Private Sub eClear_Click()
    'Menu Option To Clear The Text Boxes
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub ehide_Click()
    'Menu Option To Change Password Character To *
    Text2.PasswordChar = "*"
End Sub

Private Sub eshow_Click()
    'Menu Option To Remove Password Character
    Text2.PasswordChar = ""
End Sub

Private Sub fc_Click()
    'Menu Option To Create Folder
    Call Command3_Click
End Sub

Private Sub fcn_Click()
    'Menu Option To Change Secure Folder Name
    FormNameChange.Show
End Sub

Private Sub fd_Click()
    'Menu Option To Delete Secure Folder
    Call Command4_Click
End Sub

Private Sub fe_Click()
    'Menu Option To Exit
    End
End Sub

Private Sub fl_Click()
    'Menu Option To Lock Secure Folder
    Call Command1_Click
End Sub

Private Sub Form_Load()
    choice = 0
End Sub

Private Sub FSet_Click()
    Dim sTempDir As String, i As Integer
    i = 0
    On Error GoTo error_handle_openfile
    sTempDir = CurDir    'Remember the current active directory
    CommonDialog1.DialogTitle = "Select Directory For Secure Folder" 'titlebar
    CommonDialog1.InitDir = App.Path 'start dir, might be "C:\" or so also
    CommonDialog1.FileName = "Select Directory"  'Something in filenamebox
    CommonDialog1.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog1.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    CommonDialog1.CancelError = True 'allow escape key/cancel
    CommonDialog1.ShowOpen 'show the dialog screen
    dir = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 20)
    pass1 = write_file(App.Path + "/Folder_lock_data/dirlocation.txt", encrypt(dir + decrypt(read_file(App.Path + "/Folder_lock_data/dirname.txt"))))
    i = 1
    ChDir sTempDir  'restore path to what it was at entering
    choice = 3
If i = 0 Then
error_handle_openfile:
    choice = MsgBox("No Folder Is Selected to Place Secure Folder . Do You Want To Try Again?", vbOKCancel, "File Selection Aborted")
    If choice = 1 Then
        Call FSet_Click
    End If
End If
End Sub

Private Sub fShow_Click()
    Dim i As Integer
    i = 0
    On Error GoTo Error_Handle_Path
    MsgBox "Your Secure Folder Location Is : " + decrypt(read_file(App.Path + "/Folder_Lock_data/dirlocation.txt")), vbOKOnly, "Secure Folder Path"
    i = 1
If i = 0 Then
Error_Handle_Path:
    MsgBox "Secure Folder Was Not Created. Create Secure Folder First and Then Try Again .....!!!", vbCritical, "Incomplete Action"
End If
End Sub

Private Sub fu_Click()
    'Menu Option To Unlock Secure Folder
    Call Command2_Click
End Sub

Private Sub fusr_Click()
    'Menu Option To Change Username and Password of Secure Folder
    frmLogin.Visible = True
End Sub

Private Sub ha_Click()
    'Menu Option To Make ABOUT FORM Visible
    frmAbout.Visible = True
End Sub

Private Sub Option1_Click()
    'Option To Show Password Without Special Character
    Text2.PasswordChar = ""
End Sub

Private Sub vc_Click()
    'Option To Change Form Backcolor
    CommonDialog2.ShowColor
    Form1_MAIN.BackColor = CommonDialog2.Color
    FormNameChange.BackColor = CommonDialog2.Color
    frmLogin.BackColor = CommonDialog2.Color
End Sub
