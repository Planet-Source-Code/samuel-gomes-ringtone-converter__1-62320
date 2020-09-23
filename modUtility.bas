Attribute VB_Name = "modUtility"
' Miscellaneous utility functions
' Copyright (c) Samuel Gomes (Blade), 2003-2004
' mailto: v_2samg@hotmail.com

Option Explicit

' Global type definitions
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pIdList As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Const MAX_PATH = 260
Private Const BIF_RETURNONLYFSDIRS = &H1      ' For finding a folder to start document searching
Private Const BIF_DONTGOBELOWDOMAIN = &H2     ' For starting the Find Computer
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20             ' insist on valid result (or CANCEL)
Private Const BIF_BROWSEFORCOMPUTER = &H1000  ' Browsing for Computers.
Private Const BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
Private Const BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything

' An empty string constant, makes the code more readable :)
Public Const sEmpty = ""

' Converts a null-terminated string to a VB string (more of a convenience... trims it)
Public Function CStrToBStr(ByVal lpszString As String) As String
    lpszString = lpszString & vbNullChar
    CStrToBStr = Left(lpszString, InStr(lpszString, vbNullChar) - 1)
End Function

' Clamps vVal between vMin and vMax
Public Function Clamp(ByVal vVal As Variant, ByVal vMin As Variant, ByVal vMax As Variant) As Variant
    Clamp = IIf(vVal > vMax, vMax, IIf(vVal < vMin, vMin, vVal))
End Function

' Similar to the C library strbrk() function
Public Function StrBrk(ByVal InString As String, ByVal Separator As String) As Long
    Dim ln As Long
    Dim BegPos As Long
    
    ln = Len(InString)
    BegPos = 1

    Do While (InStr(Separator, Mid(InString, BegPos, 1)) = 0)
        If (BegPos > ln) Then
            StrBrk = 0
            Exit Function
        Else
            BegPos = BegPos + 1
        End If
    Loop

    StrBrk = BegPos
End Function

' Similar to the C library strspn() function
Public Function StrSpn(ByVal InString As String, ByVal Separator As String) As Long
    Dim ln As Long
    Dim BegPos As Long
    
    ln = Len(InString)
    BegPos = 1

    Do While (InStr(Separator, Mid(InString, BegPos, 1)) <> 0)
        If (BegPos > ln) Then
            StrSpn = 0
            Exit Function
        Else
            BegPos = BegPos + 1
        End If
    Loop

    StrSpn = BegPos
End Function

' The main string parsing workhorse
Public Function GetToken(ByVal Search As String, ByVal Delim As String) As String
    Static SaveStr As String, BegPos As Long
    Dim newPos As Long

    If (Search <> sEmpty) Then
      BegPos = 1
      SaveStr = Search
    End If

    newPos = StrSpn(Mid(SaveStr, BegPos, Len(SaveStr)), Delim)
    If (newPos) Then
        BegPos = newPos + BegPos - 1
    Else
        GetToken = sEmpty
        Exit Function
    End If

    newPos = StrBrk(Mid(SaveStr, BegPos, Len(SaveStr)), Delim)
    If (newPos) Then
        newPos = BegPos + newPos - 1
    Else
        newPos = Len(SaveStr) + 1
    End If
    GetToken = Mid(SaveStr, BegPos, newPos - BegPos)
    BegPos = newPos
End Function

' Just a convenient wrapper over GetToken()
Public Function ParseString(ByVal UserString As String, ByVal UserToken As String, ByVal SubStringNumber As Long) As String
    Dim i As Long
    
    If (UserString = sEmpty) Then
        ParseString = sEmpty
        Exit Function
    End If

    ParseString = GetToken(UserString, UserToken)
    If (SubStringNumber < 2) Then Exit Function
    i = 1
    Do
        ParseString = GetToken(sEmpty, UserToken)
        i = i + 1
    Loop While (i < SubStringNumber)
End Function

' Displays an error dialog
Public Function ErrorDialog(Optional ByVal sMessage As String = sEmpty, Optional ByVal sErrSrc As String = sEmpty, Optional ByVal sErrApp As String = sEmpty) As Integer
    If (sMessage = sEmpty) Then
        sMessage = Error
        If (sMessage = sEmpty) Then
            sMessage = "Unknown error"
        End If
    End If
    
    If (sErrApp = sEmpty) Then
        sErrApp = App.ProductName
        If (sErrApp = sEmpty) Then
            sErrApp = App.Title
        End If
    End If
    
    If (sErrSrc = sEmpty) Then
        sErrSrc = Err.Source
        If (sErrSrc = sEmpty) Then
            sErrSrc = App.EXEName & ".exe"
        End If
    End If
    
    ErrorDialog = MsgBox("The following error occured in " & sErrApp & " (" & sErrSrc & "):" & vbCrLf & vbCrLf & sMessage & "!", vbAbortRetryIgnore Or vbCritical Or vbDefaultButton2 Or vbMsgBoxSetForeground)
End Function

Public Function BrowseForFolderDialog(ByRef frmForm As Object, Optional ByVal sTitle As String = "Browse For Folder:", Optional ByVal bShowFiles As Boolean = False, Optional bShowEditBox As Boolean = False) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo

    ' Setup some data
    With tBrowseInfo
        .hWndOwner = frmForm.hwnd
        .lpszTitle = sTitle
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or IIf(bShowFiles, BIF_BROWSEINCLUDEFILES, 0) Or IIf(bShowEditBox, BIF_EDITBOX, 0)
    End With

    ' Finally...
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    ' Yuck!
    If (CBool(lpIDList)) Then
        sBuffer = Space(MAX_PATH)
        If (CBool(SHGetPathFromIDList(lpIDList, sBuffer))) Then
            BrowseForFolderDialog = CStrToBStr(sBuffer)
        End If
        CoTaskMemFree lpIDList
    End If
End Function

Function MakeLegalFileName(ByVal sFName As String) As String
    Dim i As Long

    For i = 1 To Len(sFName)
        If (InStr("\/:*?<>|" + Chr(34), Mid(sFName, i, 1)) > 0) Then
            Mid(sFName, i, 1) = "_"
        End If
    Next

    MakeLegalFileName = Trim(sFName)
End Function

