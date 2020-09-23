Attribute VB_Name = "mod_BrowseDialog"
Option Explicit

'// public win32 api declarations
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'// public constants declaration
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const MAX_PATH = 260
'// public data types declaration
Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Function SHBrowseDialog(ByVal Title As String, _
    ByVal Form As Form) As String
    
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = Title

    With tBrowseInfo
        .hwndOwner = Form.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        SHBrowseDialog = sBuffer
    End If
End Function
