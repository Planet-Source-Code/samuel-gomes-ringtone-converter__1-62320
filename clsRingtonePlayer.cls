VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRingTonePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Ringtone playback class
' Copyright (c) Samuel Gomes (Blade), 2003-2004
' mailto: v_2samg@hotmail.com

' This implements ringtone playback in the application.
' It is responsible for playing a single tone given the note, duration etc.
' It also handles silence.
' The details:
' Supports playback of the folowing notes: c, d, e, f, g, a, b and their #s
' Supports octaves: 1, 2, 3 where 1 is the lowest and 3 is the highest
' Supports dot notes: '.' (affects duration of the note played back)
' Supports durations: 1, 2, 3, 4, 5, 6 (where 1 is the shortest and 6 is the longest or full)
' Supports silence: p (p = pause :)
' Hence a playable note may look like '6.c#3'
' And a silence may look like '6p' (no dots here!)
' BTW this is an internal format we use, just to keep the playback part consistent and simple.
' It's up to the individual 'Ringtone classes' to convert it's data to this format.
' As you can see, if we do this, then interconversion between various formats becomes very simple. :)
' Also, it supports a global tempo, which ranges from 1 to 255 (200 is the default)

' If you want to implement the note output using MIDI, DirectSound or whatever, then this is the
' module to change.

' This class contains a bug: Sometimes some tone does not 'sound' right. Probably some miscalculation
' is taking place or some data is not right. NEEDS TO BE FIXED!

' Must use this!
Option Explicit

' Kernel function we'll use to listen to the ringtone (works only on Windows NT/2000/XP/2003... :()
' We do not use MIDI or DirectSound just to keep the source short and simple and I don't have the
' time to implement it. You are more than welcome to do so. Send me a copy if you do and I'll
' update the source release.
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' The main note frequencies. We calculate the sharps mathematically. :)
' I hope these are correct. Can someone provide me values with more precision?
Private Const NOTE_FRQ_C = 261.63!
Private Const NOTE_FRQ_D = 293.66!
Private Const NOTE_FRQ_E = 329.63!
Private Const NOTE_FRQ_F = 349.23!
Private Const NOTE_FRQ_G = 392!
Private Const NOTE_FRQ_A = 440!
Private Const NOTE_FRQ_B = 493.88!
Private Const NOTE_FRQ_PAUSE = 32767        ' dummy!
' Note duration and octave constants
Private Const NOTE_OCTAVE_MIN = 1
Private Const NOTE_OCTAVE_MAX = 3
' The full note duration in ms. The rest as usual is calculated mathematically. ;)
Private Const NOTE_DURATION_FULL = 1200
Private Const NOTE_DURATION_MIN = 1
Private Const NOTE_DURATION_MAX = 6
' Default tempo. Why 200? Look at the comments later :)
Private Const NOTE_TEMPO_DEFAULT = 200
Private Const NOTE_TEMPO_MIN = 1
Private Const NOTE_TEMPO_MAX = 255
' String representations of the vaid notes and symbols
Private Const NOTE_C = "C"
Private Const NOTE_D = "D"
Private Const NOTE_E = "E"
Private Const NOTE_F = "F"
Private Const NOTE_G = "G"
Private Const NOTE_A = "A"
Private Const NOTE_B = "B"
Private Const NOTE_PAUSE = "P"
Private Const NOTE_DOT = "."
Private Const NOTE_SHARP = "#"
Private Const NOTE_CHARS = NOTE_C & NOTE_D & NOTE_E & NOTE_F & NOTE_G & NOTE_A & NOTE_B & NOTE_PAUSE & NOTE_DOT & NOTE_SHARP

' Global class variables
Private cTempo As Byte
' Hold the ringtone in internal format
Private colRingtone As Collection
' Which note we are currently pointing to?
Private lNoteIndex As Long
' Ringtone optimization flag
Private bOptimize As Boolean
Private sToneName As String

' Events
Public Event Playing(ByVal sNote As String, ByVal fFrequency As Single, ByVal fDuration As Single)

Private Sub Class_Initialize()
    ' Nothing much to de here... clear up
    Clear
End Sub

Private Sub Class_Terminate()
    ' Again, not much to do here. But might be useful if we are doing MIDI or DirectSound
    Set colRingtone = Nothing
End Sub

' Calculates the duration in MS
Private Function CalcDuration(ByVal iDuration As Integer, ByVal bIsDot As Boolean) As Single
    Dim fDuration As Single
    
    ' Calculate the duration in milliseconds
    fDuration = CSng(NOTE_DURATION_FULL) / (2! ^ (NOTE_DURATION_MAX - iDuration))
    ' Correct duration if we have a dot note.
    ' A dot note is the average of this duration and the "next" duration
    If (bIsDot) Then fDuration = (fDuration * 3!) / 2!
    ' Calculate correct duration based on tempo
    ' Some history: BPM = Beats/minute = quater notes/minute
    ' In one minute we have 60 secs or 60 x 1000 = 60000 ms
    ' One quater note = 300 ms
    ' Therefore the normal BPM = 60000 / 300 = 200
    CalcDuration = fDuration * (CSng(NOTE_TEMPO_DEFAULT) / CSng(cTempo))
End Function

' Calculates the frequency
Private Function CalcFrequency(ByVal cNote As Byte, ByVal bIsSharp As Boolean, ByVal cOctave As Byte) As Single
    Dim fFrequency As Single, fTemp As Single
    
    ' Get the frequency for the corresponding note (for below middle c octave)
    fFrequency = Choose(cNote, NOTE_FRQ_C, NOTE_FRQ_D, NOTE_FRQ_E, NOTE_FRQ_F, NOTE_FRQ_G, NOTE_FRQ_A, NOTE_FRQ_B)
    ' Correct frequency if we have a sharp (i.e. mean of this note and the "next" note)
    If (bIsSharp) Then
        ' Get the frequency of next note. Note how we rotated the notes to the left (no pun intended)
        ' Handle special case 'c' from next octave
        fTemp = Choose(cNote, NOTE_FRQ_D, NOTE_FRQ_E, NOTE_FRQ_F, NOTE_FRQ_G, NOTE_FRQ_A, NOTE_FRQ_B, NOTE_FRQ_C * 2!)
        ' Calculate sharp
        fFrequency = (fFrequency + fTemp) / 2!
    End If
    ' Calculate frequency for the corresponding octave
    CalcFrequency = fFrequency * CSng(cOctave)
End Function

' Gets the current full note
Private Function GetFullNote() As String
    ' Ignore if index is out of bounds
    On Error Resume Next
    GetFullNote = colRingtone.Item(CStr(lNoteIndex))
End Function

' Tempo get set methods
Public Property Get Tempo() As Byte
    Tempo = cTempo
End Property

Public Property Let Tempo(ByVal cTmpo As Byte)
    If (cTmpo < NOTE_TEMPO_MIN) Then
        Err.Raise vbObjectError + 1011, , "Tempo out of range"
    End If
    
    cTempo = cTmpo
End Property

' Tone name get set methods
Public Property Get Name() As String
    Name = sToneName
End Property

Public Property Let Name(ByVal sName As String)
    sToneName = sName
End Property

' Sets the tone optimization flag
Public Property Get Optimize() As Boolean
    Optimize = bOptimize
End Property

Public Property Let Optimize(ByVal bOpt As Boolean)
    bOptimize = bOpt
End Property

' Clears the ringtone data
Public Sub Clear()
    cTempo = NOTE_TEMPO_DEFAULT
    Set colRingtone = Nothing
    Set colRingtone = New Collection
    lNoteIndex = 0
    sToneName = sEmpty
End Sub

' Adds a note to the ringtone
Public Sub AddNote(ByVal sNote As String, ByVal bIsSharp As Boolean, ByVal cOctave As Byte, ByVal cDuration As Byte, ByVal bIsDot As Boolean)
    Dim sFullNote As String, sPrevNote As String
    
    ' Validate all parameters
    sNote = UCase(Trim(sNote))
    
    Select Case sNote
        Case NOTE_C, NOTE_D, NOTE_E, NOTE_F, NOTE_G, NOTE_A, NOTE_B
            ' Ok
        Case NOTE_PAUSE
            If (bIsDot) Then
                Err.Raise vbObjectError + 1015, , "A pause cannot be a dot note"
            End If
        Case Else
            Err.Raise vbObjectError + 1012, , "Invalid note"
    End Select
    
    If (cOctave < NOTE_OCTAVE_MIN Or cOctave > NOTE_OCTAVE_MAX) Then
        Err.Raise vbObjectError + 1013, , "Note octave out of range"
    End If
    
    If (cDuration < NOTE_DURATION_MIN Or cDuration > NOTE_DURATION_MAX) Then
        Err.Raise vbObjectError + 1014, , "Note duration out of range"
    End If
    
    ' Create the full note
    sFullNote = cDuration
    If (sNote <> NOTE_PAUSE And bIsDot) Then sFullNote = sFullNote & NOTE_DOT
    sFullNote = sFullNote & sNote
    If (sNote <> NOTE_PAUSE And bIsSharp) Then sFullNote = sFullNote & NOTE_SHARP
    If (sNote <> NOTE_PAUSE) Then sFullNote = sFullNote & cOctave
    
    If (bOptimize And colRingtone.Count > 0) Then
        ' Club notes with duration between 1 to 5
        ' Get previous note
        sPrevNote = colRingtone.Item(CStr(colRingtone.Count - 1))
        ' If both notes are similar then proceed
        If (sPrevNote = sFullNote And cDuration < NOTE_DURATION_MAX) Then
            ' Remove the last note from the collection
            colRingtone.Remove CStr(colRingtone.Count - 1)
            ' Modify this note to reflect the duration change
            sFullNote = cDuration + 1
            If (sNote <> NOTE_PAUSE And bIsDot) Then sFullNote = sFullNote & NOTE_DOT
            sFullNote = sFullNote & sNote
            If (sNote <> NOTE_PAUSE And bIsSharp) Then sFullNote = sFullNote & NOTE_SHARP
            If (sNote <> NOTE_PAUSE) Then sFullNote = sFullNote & cOctave
        End If
    End If
    
    ' Add note to collection
    colRingtone.Add sFullNote, CStr(colRingtone.Count)
End Sub

' Moves the note index to the next note
Public Function NextNote() As Boolean
    lNoteIndex = lNoteIndex + 1
    NextNote = (GetFullNote <> sEmpty)
End Function

' Moves to the first note
Public Function FirstNote() As Boolean
    lNoteIndex = 0
    FirstNote = (GetFullNote <> sEmpty)
End Function

' Retuns the duration of the current note
Public Function GetDuration() As Byte
    GetDuration = Val(ParseString(GetFullNote, NOTE_CHARS, 1))
End Function

' Returns true if current note is a dot note
Public Function IsDot() As Boolean
    IsDot = (InStr(GetFullNote, NOTE_DOT) > 0)
End Function

' Returns the sole note from the current full note
Public Function GetNote() As String
    Dim sTemp As String
    
    sTemp = GetFullNote
    
    GetNote = Switch(InStr(sTemp, NOTE_C) > 0, NOTE_C, InStr(sTemp, NOTE_D) > 0, NOTE_D, InStr(sTemp, NOTE_E) > 0, NOTE_E, InStr(sTemp, NOTE_F) > 0, NOTE_F, InStr(sTemp, NOTE_G) > 0, NOTE_G, InStr(sTemp, NOTE_A) > 0, NOTE_A, InStr(sTemp, NOTE_B) > 0, NOTE_B, InStr(sTemp, NOTE_PAUSE) > 0, NOTE_PAUSE)
End Function

' Returns true if note is sharp
Public Function IsSharp() As Boolean
    IsSharp = (InStr(GetFullNote, NOTE_SHARP) > 0)
End Function

' Returns the current note octave
Public Function GetOctave() As Byte
    GetOctave = Val(ParseString(GetFullNote, NOTE_CHARS, 2))
End Function

' Main note playback workhorse
Public Sub Play()
    Dim sNote As String * 1
    Dim cNote As Byte
    Dim fFrequency As Single, fDuration As Single
        
    If (colRingtone.Count = 0) Then Exit Sub
    
    ' Get the sole note
    sNote = GetNote
    ' Convert note to a numeric form
    cNote = Switch(InStr(sNote, NOTE_C) > 0, 1, InStr(sNote, NOTE_D) > 0, 2, InStr(sNote, NOTE_E) > 0, 3, InStr(sNote, NOTE_F) > 0, 4, InStr(sNote, NOTE_G) > 0, 5, InStr(sNote, NOTE_A) > 0, 6, InStr(sNote, NOTE_B) > 0, 7, InStr(sNote, NOTE_PAUSE) > 0, 0)
    
    ' Calculate the appropriate duration to play
    fDuration = CalcDuration(GetDuration, IsDot)
    
    If (cNote = 0) Then
        ' Simulate silence
        RaiseEvent Playing(GetFullNote, 0, fDuration)
        Beep NOTE_FRQ_PAUSE, CLng(fDuration)
    Else
        ' Calculate the frequency to play
        fFrequency = CalcFrequency(cNote, IsSharp, GetOctave)
        RaiseEvent Playing(GetFullNote, fFrequency, fDuration)
        Beep CLng(fFrequency), CLng(fDuration)
    End If
End Sub
