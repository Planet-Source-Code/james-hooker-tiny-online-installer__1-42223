VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAoC Mapper Setup"
   ClientHeight    =   2985
   ClientLeft      =   7830
   ClientTop       =   5280
   ClientWidth     =   3990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdNewDir 
      Appearance      =   0  'Flat
      Caption         =   "New dir"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin Setup.DownLoad dl1 
      Index           =   0
      Left            =   240
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdInstall 
      Appearance      =   0  'Flat
      Caption         =   "Install"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tiny Installer ©Dgmge 2003, www.dgmge.co.uk"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Set desired path, and press Install"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim downloadNumber As Integer
Dim scriptLines As Integer
Dim validScript As Boolean
Dim scriptAction As Integer
Dim fileNum As Integer
Dim dirNum As Integer
Dim fileList As Variant
Dim dirList As Variant
Dim dirsToMakeList As Variant
Dim dirsToMakeAmt As Integer
Dim installConfig(4) As String
Dim shortCutInfo(2) As String
Dim gotAShortCut As Boolean
Dim needToBackup As Boolean

Private Sub cmdInstall_Click()

    'checking to see if there are any directorys to be made
    'if so, make them
    If dirsToMakeAmt > 0 Then
        For i = 1 To dirsToMakeAmt
            If DirExists(Dir1.Path & dirsToMakeList(i)) Then
            Else
                MkDir Dir1.Path & dirsToMakeList(i)
            End If
        Next i
    End If
    
    'disableing the drive/dir/command buttons as to avoid
    'problems during installation
    Drive1.Enabled = False
    cmdInstall.Enabled = False
    cmdNewDir.Enabled = False
    Dir1.Enabled = False
    
    lblStatus.Caption = "Downloading files, please wait (0\" & fileNum & ")"
    ProgressBar2.Max = fileNum
    downloadNumber = 1
    downloadFiles (downloadNumber)
End Sub

Private Sub downloadFiles(downloadNumber As Integer)
    Dim tempPath As String
    Dim fileName As String
    
    fileName = fileList(downloadNumber)
    Do While InStr(1, fileName, "/") > 0
        fileName = Right(fileName, (Len(fileName) - InStr(1, fileName, "/")))
    Loop
    
    Load Me.dl1(downloadNumber)
    dl1(downloadNumber).Url = fileList(downloadNumber)
            
    dl1(downloadNumber).GetFileInformation
    dl1(downloadNumber).SaveLocation = Dir1.Path & dirList(downloadNumber) & fileName
    ProgressBar1.Max = dl1(downloadNumber).FileSize
    If FileExists(dl1(downloadNumber).SaveLocation) Then
            FileCopy dl1(downloadNumber).SaveLocation, dl1(downloadNumber).SaveLocation & ".bak"
            Kill dl1(downloadNumber).SaveLocation
            needToBackup = True
    End If
    dl1(downloadNumber).DownLoad
End Sub

Private Sub dl1_DLComplete(Index As Integer)
    Unload Me.dl1(Index)
    ProgressBar2.Value = downloadNumber
    lblStatus.Caption = "Downloading files, please wait (" & downloadNumber & "\" & fileNum & ")"
    downloadNumber = downloadNumber + 1
    If downloadNumber > fileNum Then
        lblStatus.Caption = "Install complete!"
        If gotAShortCut = True Then
            If shortCutInfo(2) = "Desktop" Then
                tempPath = getSpecialFolder(&H10)
                If CreateShortcut(tempPath & "\" & shortCutInfo(0), Dir1.Path & shortCutInfo(1)) Then
                    lblStatus.Caption = "Install complete, shortcut created"
                Else
                    lblStatus.Caption = "Install complete, couldnt create shortcut!"
                End If
            Else
                If CreateShortcut(shortCutInfo(2) & shortCutInfo(0), Dir1.Path & shortCutInfo(1)) Then
                    lblStatus.Caption = "Install complete, shortcut created"
                Else
                    lblStatus.Caption = "Install complete, couldnt create shortcut!"
                End If
            End If
        End If
    Else
        downloadFiles (downloadNumber)
    End If
End Sub

Private Sub dl1_DLError(Index As Integer, lpErrorDescription As String)
    MsgBox lpErrorDescription
End Sub

Private Sub dl1_RecievedBytes(Index As Integer, lnumBYTES As Long)
    On Error Resume Next
    ProgressBar1.Value = lnumBYTES
    lblSize.Caption = lnumBYTES & "/" & dl1(Index).FileSize
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub cmdNewDir_Click()
    Dim newDirName As String
    newDirName = InputBox("Dir Name", "New Dir Name")
    If Len(newDirName) > 0 Then
        MkDir Dir1.Path & "\" & newDirName
        Dir1.Path = Dir1.Path & "\" & newDirName
    End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    Dim scriptCommand As String

    'Reading the config file
    If FileExists("install.ini") Then
    Else
        MsgBox "Error, install script not found!"
        End
    End If
    Open "install.ini" For Input As #1
    While Not EOF(1)
        Line Input #1, temp$
        scriptLines = scriptLines + 1
        If Len(temp$) > 0 Then
            If Mid(temp$, 1, 1) = ":" Then
                scriptCommand = Right(temp$, (Len(temp$) - 1))
                If scriptCommand = "SCRIPTSTART" Then validScript = 1
                If scriptCommand = "INSTALLINFO" Then scriptAction = 1
                If scriptCommand = "FILELIST" Then scriptAction = 2
                If scriptCommand = "DIRSETUP" Then scriptAction = 3
                If scriptCommand = "SHORTCUT" Then scriptAction = 4
            Else
                If validScript <> True Then
                    MsgBox "Invalid install script, exiting"
                    End
                End If
                If Mid(temp$, 1, 1) = "#" Then GoTo Comment
                Select Case scriptAction
                Case 1
                    tempInt = tempInt + 1
                    scriptCommand = Mid(temp$, 1, (InStr(1, temp$, ":: ") - 1))
                    installConfig((tempInt - 1)) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                    If scriptCommand = "FILES" Then
                        ReDim fileList(1 To installConfig(4)) As String
                        ReDim dirList(1 To installConfig(4)) As String
                    End If
                Case 2
                    scriptCommand = Mid(temp$, 1, (InStr(1, temp$, ":: ") - 1))
                    If scriptCommand = "FILE" Then
                        fileNum = fileNum + 1
                        fileList(fileNum) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                    End If
                    If scriptCommand = "DIR" Then
                        dirNum = dirNum + 1
                        dirList(dirNum) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                    End If
                Case 3
                    scriptCommand = Mid(temp$, 1, (InStr(1, temp$, ":: ") - 1))
                    If scriptCommand = "AMMOUNT" Then
                        dirsToMakeAmt = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                        ReDim dirsToMakeList(1 To dirsToMakeAmt) As String
                    End If
                    If scriptCommand = "MAKE" Then
                        tempnum = tempnum + 1
                        dirsToMakeList(tempnum) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                    End If
                Case 4
                    gotAShortCut = True
                    scriptCommand = Mid(temp$, 1, (InStr(1, temp$, ":: ") - 1))
                    If scriptCommand = "NAME" Then shortCutInfo(0) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                    If scriptCommand = "PATH" Then shortCutInfo(1) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                    If scriptCommand = "LOCATION" Then shortCutInfo(2) = Right(temp$, (Len(temp$) - InStr(1, temp$, ":: ") - 2))
                End Select
            End If
        End If
Comment:
    Wend
    Close #1
    Form1.Caption = installConfig(0)
    MsgBox installConfig(3) & vbCrLf & vbCrLf & installConfig(1) & vbCrLf & installConfig(2)
End Sub

Private Sub Label1_Click()
    ShellExecute hWnd, "open", "http://www.dgmge.co.uk", vbNullString, vbNullString, conSwNormal
End Sub
