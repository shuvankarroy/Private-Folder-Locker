Attribute VB_Name = "Module1"
Public Function encrypt(s As String) As String
    Dim total As String, tmp As String, v As Long
    v = 28
    For i = 1 To Len(s)
        tmp = Mid(s, i, 1)
        tmp = Asc(tmp) + v
        tmp = Chr(tmp)
        total = total & tmp
    Next i
    encrypt = total
End Function

Public Function decrypt(s As String) As String
    Dim total As String, tmp As String, v As Long
    v = 28
    For i = 1 To Len(s)
        tmp = Mid(s, i, 1)
        tmp = Asc(tmp) - v
        tmp = Chr(tmp)
        total = total & tmp
    Next i
    decrypt = total
End Function

Public Function read_file(source As String) As String
    Open source For Input As #1
    Do While Not EOF(1)
        Line Input #1, MyLine
    Loop
    Close #1
    read_file = Mid(MyLine, 2, (Len(MyLine) - 2))
End Function

Public Function write_file(source As String, data As String) As Integer
    Open source For Output As #1
    Write #1, data
    Close #1
    write_file = 1
End Function

Function check(a As String, b As String) As Integer
    Dim x As String, y As String, MyLine As String
    x = decrypt(read_file(App.Path + "/Folder_lock_data/Username.txt"))
    y = decrypt(read_file(App.Path + "/Folder_lock_data/Password.txt"))
    If ((a = x) And (b = y)) Then
        check = 1
    Else
        check = 0
    End If
End Function

