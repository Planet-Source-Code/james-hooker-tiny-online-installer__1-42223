Attribute VB_Name = "Module1"
Private Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type
Private Type ITEMIDLIST
    shellID As SHITEMID
End Type

Const SF_DESKTOP = &H0
Const SF_PROGRAMS = &H2
Const SF_MYDOCS = &H5
Const SF_FAVORITES = &H6     ' 98+
Const SF_STARTUP = &H7
Const SF_RECENT = &H8
Const SF_SENDTO = &H9
Const SF_STARTMENU = &HB
Const SF_MYMUSIC = &HD       ' Me+
Const SF_DESKTOP2 = &H10
Const SF_NETHOOD = &H13
Const SF_FONTS = &H14
Const SF_SHELLNEW = &H15
Const SF_STARTUP2 = &H18
Const SF_ALLUSERSDESK = &H19
Const SF_APPDATA = &H1A
Const SF_PRINTHOOD = &H1B
Const SF_APPDATA2 = &H1C
Const SF_TEMPINETFILES = &H20
Const SF_COOKIES = &H21
Const SF_HISTORY = &H22
Const SF_ALLUSERSAPPDATA = &H23
Const SF_WINDOWS = &H24
Const SF_WINSYSTEM = &H25
Const SF_PROGFILES = &H26
Const SF_MYPICS = &H27       ' Me+
Const SF_USERDIR = &H28
Const SF_WINSYSTEM2 = &H29
Const SF_COMMON = &H2B

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long

Public Function getSpecialFolder(whichFolder As Long) As String
    Dim Path As String * 256
    Dim myid As ITEMIDLIST
    Dim rval As Long
    If IsMissing(useForm) Then
    rval = SHGetSpecialFolderLocation(Form1.hwnd, whichFolder, myid)
    Else
    rval = SHGetSpecialFolderLocation(Form1.hwnd, whichFolder, myid)
    End If
    
    If rval = 0 Then ' If success
      rval = SHGetPathFromIDList(ByVal myid.shellID.SHItem, ByVal Path)
        If rval Then ' If True
        getSpecialFolder = Left(Path, InStr(Path, Chr(0)) - 1)
        End If
    End If
End Function

Public Function Parse(sIn As String, sDel As String) As Variant
    Dim i As Integer, x As Integer, s As Integer, t As Integer
    i = 1: s = 1: t = 1: x = 1
    ReDim tArr(1 To x) As Variant

    If InStr(1, sIn, sDel) <> 0 Then


        Do
            ReDim Preserve tArr(1 To x) As Variant
            tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            x = x + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(1 To x) As Variant
        tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(1) = sIn
    End If
    Parse = tArr
End Function
