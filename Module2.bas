Attribute VB_Name = "Module2"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum ShortcutWindowState
    wsNormal = 1
    wsMaximized = 3
    wsMinimized = 7
End Enum

Function CreateShortcut(ByRef FullPathOfShortcut As String, DestPath As String, _
        Optional AskForReplace As Boolean = True, _
        Optional sArguments As String, _
        Optional sDescription As String, _
        Optional sHotkey As String, _
        Optional sWindowState As ShortcutWindowState = wsNormal, _
        Optional sWorkingDirectory As String, _
        Optional sIconLocation As String, _
        Optional nIconLocationIndex As Long = 0 _
        ) As Boolean
Dim oShell As New Shell, f As Folder, fItem As FolderItem
Dim shellLink As ShellLinkObject, NewLinkName As String
On Error GoTo ErrHandle1

' Add the .lnk extension if not there
If UCase(Right(FullPathOfShortcut, 3)) <> "LNK" Then _
        FullPathOfShortcut = FullPathOfShortcut & ".lnk"
        
' Lets create a link on fly
If Not CreateALink(FullPathOfShortcut, AskForReplace) Then Exit Function

'Get special folder My Computer
Set f = oShell.NameSpace(ssfDRIVES)

'Get the shortcut
Set fItem = f.ParseName(FullPathOfShortcut)
If Not fItem Is Nothing Then ' if all Ok
    Set shellLink = fItem.GetLink
    
    'set the new path and other settings
    With shellLink
        .Path = DestPath
        .Arguments = sArguments
        .Description = sDescription
        If sHotkey <> "" Then shellLink.Hotkey = sHotkey
        .ShowCommand = sWindowState
        .WorkingDirectory = sWorkingDirectory
        If sIconLocation <> "" Then _
            .SetIconLocation sIconLocation, nIconLocationIndex
    
        ' Save the changes
        shellLink.Save
    End With
    
    CreateShortcut = True
End If

Exit Function

ErrHandle1:
If Dir(FullPathOfShortcut) <> "" Then Kill FullPathOfShortcut
Exit Function

End Function

Private Function CreateALink(PathOfLink As String, Optional ReplaceConfirm As Boolean = False) As Boolean
Dim s As String
On Error GoTo ErrHandle1

If Dir(PathOfLink) <> "" Then
    If ReplaceConfirm Then
        Beep
        If MsgBox("The link file: """ & PathOfLink & """ exists. Replace it?", _
                vbYesNo) = vbNo Then Exit Function
    End If
    Kill PathOfLink
End If

' This is a shortcut file to c:\boot.ini
s = Chr(76) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(20) & Chr(2) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(192) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(70) & Chr(27) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(96) & Chr(141) & Chr(163) & Chr(217) & Chr(154) & Chr(190) & Chr(193) & Chr(1) & Chr(0) & Chr(176) & Chr(75) & Chr(24) & Chr(145) & Chr(196) & Chr(193) & Chr(1) & Chr(0) & Chr(168) & Chr(235) & Chr(178) & Chr(70) & Chr(195) & Chr(193) & Chr(1) & Chr(33) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(79) & Chr(0) & _
    Chr(20) & Chr(0) & Chr(31) & Chr(15) & Chr(224) & Chr(79) & Chr(208) & Chr(32) & Chr(234) & Chr(58) & Chr(105) & Chr(16) & Chr(162) & Chr(216) & Chr(8) & Chr(0) & Chr(43) & Chr(48) & Chr(48) & Chr(157) & Chr(25) & Chr(0) & Chr(35) & Chr(67) & Chr(58) & Chr(92) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(49) & Chr(170) & Chr(32) & Chr(0) & Chr(50) & Chr(0) & Chr(33) & Chr(1) & Chr(0) & Chr(0) & Chr(100) & Chr(44) & Chr(92) & Chr(52) & Chr(0) & Chr(0) & Chr(98) & Chr(111) & Chr(111) & Chr(116) & Chr(46) & Chr(105) & Chr(110) & Chr(105) & Chr(0) & Chr(66) & Chr(79) & Chr(79) & Chr(84) & Chr(46) & Chr(73) & Chr(78) & Chr(73) & _
    Chr(0) & Chr(0) & Chr(0) & Chr(97) & Chr(0) & Chr(0) & Chr(0) & Chr(28) & Chr(0) & Chr(0) & Chr(0) & Chr(3) & Chr(0) & Chr(0) & Chr(0) & Chr(28) & Chr(0) & Chr(0) & Chr(0) & Chr(52) & Chr(0) & Chr(0) & Chr(0) & Chr(56) & Chr(0) & Chr(0) & Chr(0) & Chr(88) & Chr(0) & Chr(0) & Chr(0) & Chr(24) & Chr(0) & Chr(0) & Chr(0) & Chr(3) & Chr(0) & Chr(0) & Chr(0) & Chr(220) & Chr(19) & Chr(40) & Chr(14) & Chr(16) & Chr(0) & Chr(0) & Chr(0) & Chr(68) & Chr(82) & Chr(73) & Chr(86) & Chr(69) & Chr(95) & Chr(67) & Chr(0) & Chr(67) & Chr(58) & Chr(92) & Chr(0) & Chr(32) & Chr(0) & Chr(0) & Chr(0) & Chr(2) & Chr(0) & Chr(0) & Chr(0) & Chr(20) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(2) & Chr(0) & Chr(92) & _
    Chr(92) & Chr(78) & Chr(73) & Chr(75) & Chr(79) & Chr(83) & Chr(95) & Chr(72) & Chr(92) & Chr(67) & Chr(0) & Chr(98) & Chr(111) & Chr(111) & Chr(116) & Chr(46) & Chr(105) & Chr(110) & Chr(105) & Chr(0) & Chr(10) & Chr(0) & Chr(46) & Chr(92) & Chr(98) & Chr(111) & Chr(111) & Chr(116) & Chr(46) & Chr(105) & Chr(110) & Chr(105) & Chr(3) & Chr(0) & Chr(67) & Chr(58) & Chr(92) & Chr(0) & Chr(0) & Chr(0) & Chr(0)

Open PathOfLink For Binary As #1
    Put #1, , s
Close #1

CreateALink = True

Exit Function

ErrHandle1:
If Dir(PathOfLink) <> "" Then Kill PathOfLink
Exit Function

End Function

Public Function FileExists(strPath As String) As Boolean
    strPath = Trim(strPath)
    If strPath = "" Then
        FileExists = False
        Exit Function
    End If
  FileExists = Len(Dir(strPath)) <> 0
End Function

Function DirExists(ByVal DName As String) As Boolean
    Dim sDummy As String
    On Error Resume Next

    If Right(DName, 1) <> "\" Then DName = DName & "\"
    sDummy = Dir$(DName & "*.*", vbDirectory)
    DirExists = Not (sDummy = "")
End Function

