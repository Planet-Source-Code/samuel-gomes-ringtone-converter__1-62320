VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRingtoneSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' iMelody format encoder
' Copyright (c) Samuel Gomes (Blade), 2003-2004
' mailto: v_2samg@hotmail.com
'
' Sony Ericsson monophonic rintones are actually iMelody (.imy) files with the following format:
'
' BEGIN:IMELODY
' VERSION:<version>
' FORMAT:<format>
' [NAME:<tone-name>]
' [COMPOSER:<compose-information>]
' [BEAT:<bpm>]
' [STYLE:<style>]
' [VOLUME:<volume>]
' MELODY:<melody>
' END:IMELODY
'
' version := <1.2>
' format := <CLASS1.0>
' name := <"sting value">
' composer := <"string value">
' beat := <25> | ... | <900>
' style := <S0> | <S1> | <S2>
' volume-modifier := <V+> | <V->
' volume := <V1> | ... | <V15> | [volume-modifier]
' basic-note := <c> | <d> | <e> | <f> | <g> | <a> | <b>
' ess-note := <&c> | <&d> | <&e> | <&f> | <&g> | <&a> | <&b>
' iss-note := <#c> | <#d> | <#e> | <#f> | <#g> | <#a> | <#b>
' basic-ess-iss-note := <basic-note> | <ess-note> | <iss-note>
' octave-prefix := <*0> | ... | <*8>
' duration := <0> | ... | <5>
' duration-specifier := <.> | <:> | <;>
' rest := <r>
' led := <ledoff> | <ledon>
' vibe := <vibeoff> | <vibeon>
' backlight := <backoff> | <backon>
' note := [octave-prefix]<basic-ess-iss-note><duration>[duration-specifier]
' silence := <rest><duration>[duration-specifier]
' repeat-count := <0> | ...         {where 0 means forever}
' repeat := "("{<silence> | <note> | <led> | <vibe> | <volume> | <backlight>}+ "@"<repeat-count> [<volume-modifier>]")"
' melody := {<silence> | <note> | <led> | <vibe> | <volume> | <backlight>}+
'
' Duration details:
'   0 = full note
'   1 = 1/2 note
'   2 = 1/4 note
'   3 = 1/8 note
'   4 = 1/16 note
'   5 = 1/32 note
' Duration modifier details:
'   . = dot note
'   : = double dot note
'   ; = 2/3 length
' Style details:
'   S0 = natural; very short rests
'   S1 = continious; no rests
'   S2 = staccato; notes shortened and rests inserted
'
' BTW I am a proud owner of a Sony Ericsson T610 mobile :)

Option Explicit

' Maximum number of options
Private Const MAX_OPTIONS = 6
' Option emumeration
Private Enum IMYOptions
    Style = 1
    Repeat
    AutoVibration
    AutoBacklight
    AutoLED
    Volume
End Enum

Private sRingtone As String, sName As String
Private cOptions(1 To MAX_OPTIONS) As Byte

Public Function GetData() As String
    GetData = sRingtone
End Function

Public Sub SetOptions(ParamArray vpaArgs() As Variant)
    Dim cOption As Byte, vVal As Variant
    
    For Each vVal In vpaArgs
        cOption = cOption + 1
        cOptions(cOption) = vVal
    Next
End Sub

' Returns the ringtone name
Public Property Get Name() As String
    Name = sName
End Property

Public Sub ConvertFrom(ByRef RTP As clsRingTonePlayer)
    Dim sNote As String * 1
    Dim lNoteCtr As Long
    Dim bVOn As Boolean, bLOn As Boolean, bBOn As Boolean
    
    ' Store name
    sName = RTP.Name
    
    sRingtone = "BEGIN:IMELODY" & vbCrLf
    sRingtone = sRingtone & "VERSION:1.2" & vbCrLf
    sRingtone = sRingtone & "FORMAT:CLASS1.0" & vbCrLf
    sRingtone = sRingtone & "NAME:" & RTP.Name & vbCrLf
    sRingtone = sRingtone & "COMPOSER:Samuel Gomes (Blade) (c) " & Year(Now) & " for Sony Ericsson Mobile" & vbCrLf
    sRingtone = sRingtone & "BEAT:" & RTP.Tempo & vbCrLf
    sRingtone = sRingtone & "STYLE:S" & Clamp(cOptions(Style), 0, 2) & vbCrLf
    sRingtone = sRingtone & "VOLUME:V" & Clamp(cOptions(Volume), 1, 15) & vbCrLf
    sRingtone = sRingtone & "MELODY:"
    
    If (cOptions(Repeat) > 0) Then
        sRingtone = sRingtone & "(r1"
    Else
        sRingtone = sRingtone & "r5"
    End If
    
    If RTP.FirstNote Then
        Do
            sNote = LCase(RTP.GetNote)
            If (sNote = "p") Then
                sRingtone = sRingtone & "r" & 6 - RTP.GetDuration
            Else
                sRingtone = sRingtone & "*" & 3 + RTP.GetOctave & IIf(RTP.IsSharp, "#", sEmpty) & sNote & 6 - RTP.GetDuration & IIf(RTP.IsDot, ".", sEmpty)
            End If
            
            ' Increase note counter
            lNoteCtr = lNoteCtr + 1
            
            ' Backlight effect
            If (cOptions(AutoBacklight) > 0) Then
                If (lNoteCtr Mod cOptions(AutoBacklight) = 0) Then
                    bBOn = Not bBOn     ' flip state
                    sRingtone = sRingtone & IIf(bBOn, "backon", "backoff")
                End If
            End If
            
            ' LED effect
            If (cOptions(AutoLED) > 0) Then
                If (lNoteCtr Mod cOptions(AutoLED) = 0) Then
                    bLOn = Not bLOn     ' flip state
                    sRingtone = sRingtone & IIf(bLOn, "ledon", "ledoff")
                End If
            End If
            
            ' Vibration effect
            If (cOptions(AutoVibration) > 0) Then
                If (lNoteCtr Mod cOptions(AutoVibration) = 0) Then
                    bVOn = Not bVOn     ' flip state
                    sRingtone = sRingtone & IIf(bVOn, "vibeon", "vibeoff")
                End If
            End If
        Loop While RTP.NextNote
    End If
    
    ' Turn on backlight if it is not active
    If (cOptions(AutoBacklight) > 0) Then
        If (Not bBOn) Then
            sRingtone = sRingtone & "backon"
        End If
    End If
    
    ' Turn off LED if it is active
    If (cOptions(AutoLED) > 0) Then
        If (bLOn) Then
            sRingtone = sRingtone & "ledoff"
        End If
    End If
    
    ' Turn off vibration if it is active
    If (cOptions(AutoVibration) > 0) Then
        If (bVOn) Then
            sRingtone = sRingtone & "vibeoff"
        End If
    End If
    
    If (cOptions(Repeat) > 0) Then
        sRingtone = sRingtone & "@" & cOptions(Repeat) & ")"
    Else
        sRingtone = sRingtone & "r5"
    End If
    
    sRingtone = sRingtone & vbCrLf & "END:IMELODY"
End Sub

Private Sub Class_Initialize()
    sRingtone = sEmpty
    cOptions(AutoVibration) = 8
    cOptions(AutoLED) = 2
    cOptions(AutoBacklight) = 1
    cOptions(Repeat) = 4
    cOptions(Style) = 1
    cOptions(Volume) = 15
End Sub
