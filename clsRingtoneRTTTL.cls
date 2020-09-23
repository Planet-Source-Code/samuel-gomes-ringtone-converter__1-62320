VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRingtoneRTTTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Standard RTTTL format to the internal ringtone format converter
' Copyright (c) Samuel Gomes (Blade), 2003-2004
' mailto: v_2samg@hotmail.com

Option Explicit

Private sRingtone As String

' Stores the ringtone in a internal variable
Public Property Let Data(ByVal sRT As String)
    sRingtone = sRT
End Property

Public Property Get Data() As String
    Data = sRingtone
End Property

' Returns the ringtone name
Public Property Get Name() As String
    Name = Trim(ParseString(sRingtone, ":", 1))
End Property

' Converts the RTTTL ringtone to the internal ringtone format
Public Sub ConvertTo(ByRef RTP As clsRingTonePlayer)
    Dim i As Long, sParams As String, sTemp As String
    Dim cDefDuration As Byte, cDefOctave As Byte, cDefTempo As Byte
    Dim sData As String, cDuration As Byte, bIsDot As Boolean
    Dim sNote As String * 1, cOctave As Byte, bIsSharp As Boolean
    Dim iUser As Integer
    
    ' Start converting to internal format
    On Error GoTo errConvert
    
    ' Parse default parameters
    cDefDuration = 4
    cDefOctave = 5
    cDefTempo = 200
    sParams = ParseString(sRingtone, ":", 2)
    
    For i = 1 To 3
        sTemp = LCase(ParseString(sParams, ", " & vbTab, i))
        
        If (InStr(sTemp, "o") > 0) Then
            cDefOctave = Val(ParseString(sTemp, "= " & vbTab, 2))
        End If
        
        If (InStr(sTemp, "b") > 0) Then
            cDefTempo = Val(ParseString(sTemp, "= " & vbTab, 2))
        End If
        
        If (InStr(sTemp, "d") > 0) Then
            cDefDuration = Val(ParseString(sTemp, "= " & vbTab, 2))
        End If
    Next
    
    ' Initialize player
    RTP.Clear
    RTP.Name = Name
    RTP.Tempo = cDefTempo
    
    sData = ParseString(sRingtone, ":", 3)
    i = 0
    
    Do
        i = i + 1
        sTemp = LCase(ParseString(sData, ", " & vbTab, i))
        If (sTemp = sEmpty) Then Exit Do
        If (Val(sTemp) < 1) Then sTemp = "0" & sTemp     ' zero duration atleast; helps to parse
        
        sNote = "p"     ' default note
        sNote = Switch(InStr(sTemp, "c") > 0, "c", InStr(sTemp, "d") > 0, "d", InStr(sTemp, "e") > 0, "e", InStr(sTemp, "f") > 0, "f", InStr(sTemp, "g") > 0, "g", InStr(sTemp, "a") > 0, "a", InStr(sTemp, "b") > 0, "b", InStr(sTemp, "p") > 0, "p")
        bIsSharp = (InStr(sTemp, "#") > 0)
        cOctave = Val(ParseString(sTemp, "cdefgabp.#", 2))
        cOctave = IIf(cOctave < 1, cDefOctave, cOctave) - 4
        cOctave = Clamp(cOctave, 1, 3)
        cDuration = Val(ParseString(sTemp, "cdefgabp.#", 1))
        cDuration = IIf(cDuration < 1, cDefDuration, cDuration)
        cDuration = Switch(cDuration > 16, 1, cDuration > 8, 2, cDuration > 4, 3, cDuration > 2, 4, cDuration > 1, 5, cDuration >= 0, 6)
        bIsDot = (InStr(sTemp, ".") > 0 And sNote <> "p")
        
        RTP.AddNote sNote, bIsSharp, cOctave, cDuration, bIsDot
    Loop While (sTemp <> sEmpty)
    
    Exit Sub

errConvert:
    iUser = ErrorDialog(Error & " when parsing (" & sTemp & ")")
    If (iUser = vbRetry) Then
        Resume
    ElseIf (iUser = vbIgnore) Then
        Resume Next
    End If
End Sub
