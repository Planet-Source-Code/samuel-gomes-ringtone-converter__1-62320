VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ringtone Converter"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSE 
      Caption         =   "&Sony Ericsson Format:"
      Height          =   2295
      Left            =   113
      TabIndex        =   5
      Top             =   3000
      Width           =   7455
      Begin VB.TextBox txtDestination 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         ToolTipText     =   "Displays the ringtone in the output format after conversion."
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame fraRTTL 
      Caption         =   "&RTTL Format:"
      Height          =   2295
      Left            =   113
      TabIndex        =   2
      Top             =   480
      Width           =   7455
      Begin VB.TextBox txtSource 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "Displays the ringtone in RTTTL format."
         Top             =   600
         Width           =   7215
      End
      Begin VB.ComboBox cboRingtones 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   120
         List            =   "frmMain.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select the ringtone you want."
         Top             =   240
         Width           =   7215
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5385
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Ready!"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Creates a new RTTL."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Saves the entire RTTL collection."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remove"
            Object.ToolTipText     =   "Removes ringtone from the RTTL collection."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Shows RTTL properties."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert"
            Object.ToolTipText     =   "Converts ringtone to the IMelody (IMY) format."
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Object.ToolTipText     =   "Exports ringtone to an IMelody (IMY) file."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Configures application options."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Play"
            Object.ToolTipText     =   "Plays the selected ringtone."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stops the current ringtone playback."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Shows application help information."
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6473
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0446
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0558
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E32
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1284
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1396
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A8
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15BA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A0C
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B1E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C30
            Key             =   "Convert"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRemove 
         Caption         =   "&Remove..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileConvert 
         Caption         =   "&Convert"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "&Play"
      Begin VB.Menu mnuPlayPlay 
         Caption         =   "&Play"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPlayStop 
         Caption         =   "&Stop"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Ringtone Converter..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ringtone Converter (SE Edition)
' Copyright (c) Samuel Gomes (Blade), 2003-2005
' mailto: v_2samg@hotmail.com

' Notes:
' The application only generates Nokia composer format (3315 compatible) and Panasonic composer
' format (GD68 compatible) ringtones from standard RTTTL format (as of now). Export to more formats
' may be added later.
' Ok, this is old. :( It only supports Sony Ericsson IMY format now. :)
' BTW, I love Sony Ericsson! :)

' I am a hardcore C++ programmer. Hope you'll understand.
Option Explicit

' A quick and dirty to way to implement an About dialog
Private Declare Function ShellAbout Lib "shell32" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

' The RTTTL ringtone collection file
Private Const FileRTC = "\ringtones.rtc"

' In-memory ringtone collection
Private colRingtones As New Collection
' Our global Ringtone player object
Public WithEvents rtPlayer As clsRingTonePlayer
Attribute rtPlayer.VB_VarHelpID = -1
' Our global RTTTL object
Private rtRTTTL As New clsRingtoneRTTTL
' Out global SE Ringtone converter object
Public rtSE As New clsRingtoneSE

' Converts ringtone data from RTTL to SE IMY
Private Sub mnuFileConvert_Click()
    rtRTTTL.Data = txtSource.Text
    rtRTTTL.ConvertTo rtPlayer
    rtSE.ConvertFrom rtPlayer
    txtDestination.Text = rtSE.GetData
End Sub

Private Sub mnuFileOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShellAbout Me.hwnd, App.Title, App.LegalCopyright & vbCrLf & "Version " & App.Major & "." & App.Minor & "." & App.Revision, Me.Icon
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileProperties_Click()
    MsgBox "Ringtone: " & rtRTTTL.Name, vbInformation
End Sub

Private Sub cboRingtones_Click()
    txtSource.Text = colRingtones.Item(CStr(cboRingtones.ItemData(cboRingtones.ListIndex)))
    rtRTTTL.Data = txtSource.Text
    rtRTTTL.ConvertTo rtPlayer
End Sub

Private Sub mnuFileRemove_Click()
    If (MsgBox("Are you sure that you want to remove this ringtone?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes) Then
        colRingtones.Remove CStr(cboRingtones.ItemData(cboRingtones.ListIndex))
        cboRingtones.RemoveItem cboRingtones.ListIndex
        
        If (cboRingtones.ListCount > 0) Then cboRingtones.ListIndex = 0
        
        EnableDisableControls
    End If
End Sub

Private Sub mnuFileExport_Click()
    Dim sExportFile As String
    Dim iFileExport As Integer
    
    sExportFile = BrowseForFolderDialog(Me, "Select location to export '" & rtSE.Name & "':")
    If (sExportFile = sEmpty) Then Exit Sub
    If (Right(sExportFile, 1) <> "\") Then sExportFile = sExportFile & "\"
    sExportFile = sExportFile & MakeLegalFileName(rtSE.Name) & ".imy"
    
    On Error GoTo errExport
    
    iFileExport = FreeFile
    Open sExportFile For Output Access Write As iFileExport
    
    Print #iFileExport, txtDestination.Text
    
    Close iFileExport

    Exit Sub
    
errExport:
    MsgBox "Failed to export ringtone to file (" & sExportFile & ")!", vbExclamation
End Sub

Private Sub mnuFileNew_Click()
    Dim sTitle As String
    Dim lCtr As Long
    
    ' Get title from user
    sTitle = Trim(InputBox("Enter the ringtone name:", , "Untitled"))
    
    If (sTitle = sEmpty) Then Exit Sub
    
    txtSource.Text = sTitle & ":d=4,o=5,b=200:"
    
    ' Add the ringtone to the internal list
    lCtr = cboRingtones.ListCount
    cboRingtones.AddItem sTitle, lCtr
    cboRingtones.ItemData(lCtr) = lCtr
    colRingtones.Add txtSource.Text, CStr(lCtr)
    
    cboRingtones.ListIndex = cboRingtones.ListCount - 1
    
    EnableDisableControls
End Sub

Private Sub mnuFileSave_Click()
    Dim iFileRTC As Integer
    Dim lCtr As Long
    Dim sRT As String
    
    If (MsgBox("Are you sure that you want to overwrite the entire ringtone collection?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes) Then
        ' First update the current ringtone
        colRingtones.Remove CStr(cboRingtones.ItemData(cboRingtones.ListIndex))
        colRingtones.Add txtSource.Text, CStr(cboRingtones.ItemData(cboRingtones.ListIndex))
        
        On Error GoTo errWriteRTC
        
        iFileRTC = FreeFile
        Open App.Path & FileRTC For Output Access Write As iFileRTC
        
        For lCtr = 0 To cboRingtones.ListCount - 1
            If (LCase(Trim(sRT)) <> LCase(Trim(colRingtones.Item(CStr(cboRingtones.ItemData(lCtr)))))) Then
                Print #iFileRTC, Trim(colRingtones.Item(CStr(cboRingtones.ItemData(lCtr))))
            End If
            sRT = colRingtones.Item(CStr(cboRingtones.ItemData(lCtr)))
        Next
        
        Close iFileRTC
    End If
    
    Exit Sub

errWriteRTC:
    MsgBox "Failed to save ringtones to the ringtone collection file (" & App.Path & FileRTC & ")!", vbExclamation
End Sub

Private Sub Form_Load()
    Dim iFileRTC As Integer
    Dim sRingtone As String
    Dim lCtr As Long
    
    ' Initialize the randomizer
    Randomize
    
    ' Warn the user about Windows 9x
    If (Environ("OS") <> "Windows_NT") Then
        MsgBox "You seem to be using a Windows 9x OS. The 'Play' function might not work on this OS!", vbExclamation
    End If
    
    ' Create the player objects
    Set rtPlayer = New clsRingTonePlayer
    
    ' Load all ringtones from the ringtone file into the combo
    On Error GoTo errNoRTC
    
    iFileRTC = FreeFile
    Open App.Path & FileRTC For Input Access Read As iFileRTC
    
    Do Until EOF(iFileRTC)
        Line Input #iFileRTC, sRingtone
        cboRingtones.AddItem Trim(ParseString(sRingtone, ":", 1)), lCtr
        cboRingtones.ItemData(lCtr) = lCtr
        colRingtones.Add sRingtone, CStr(lCtr)
        lCtr = lCtr + 1
    Loop
    
    Close iFileRTC
    
    cboRingtones.ListIndex = 0
    txtDestination_Change
    
    EnableDisableControls
    
    Exit Sub
    
errNoRTC:
    MsgBox "Failed to load ringtones from the ringtone collection file (" & App.Path & FileRTC & ")!", vbExclamation
End Sub

Private Sub EnableDisableControls()
    cboRingtones.Enabled = (cboRingtones.ListCount > 0)
    txtSource.Enabled = (cboRingtones.ListCount > 0)
    txtDestination.Enabled = (cboRingtones.ListCount > 0)
    mnuFileConvert.Enabled = (cboRingtones.ListCount > 0)
    mnuPlayPlay.Enabled = (cboRingtones.ListCount > 0)
    mnuFileRemove.Enabled = (cboRingtones.ListCount > 0)
    mnuFileSave.Enabled = (cboRingtones.ListCount > 0)
    mnuPlayStop.Enabled = False
    mnuFileProperties.Enabled = (cboRingtones.ListCount > 0)
    
    tbToolbar.Buttons("Remove").Enabled = mnuFileRemove.Enabled
    tbToolbar.Buttons("Save").Enabled = mnuFileSave.Enabled
    tbToolbar.Buttons("Convert").Enabled = mnuFileConvert.Enabled
    tbToolbar.Buttons("Play").Enabled = mnuPlayPlay.Enabled
    tbToolbar.Buttons("Stop").Enabled = mnuPlayStop.Enabled
    tbToolbar.Buttons("Properties").Enabled = mnuFileProperties.Enabled
    
    sbStatus.SimpleText = "Ready!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not mnuPlayPlay.Enabled
End Sub

Private Sub mnuPlayPlay_Click()
    mnuPlayPlay.Enabled = False
    tbToolbar.Buttons("Play").Enabled = mnuPlayPlay.Enabled
    mnuPlayStop.Enabled = True
    tbToolbar.Buttons("Stop").Enabled = mnuPlayStop.Enabled
    
    If rtPlayer.FirstNote Then
        Do
            rtPlayer.Play
            DoEvents
        Loop While rtPlayer.NextNote And mnuPlayStop.Enabled = True
    End If
    
    EnableDisableControls
End Sub

Private Sub mnuPlayStop_Click()
    mnuPlayStop.Enabled = False
End Sub

Private Sub rtPlayer_Playing(ByVal sNote As String, ByVal fFrequency As Single, ByVal fDuration As Single)
    If (fFrequency = 0) Then
        sbStatus.SimpleText = "Playing (" & sNote & ") silence for " & fDuration & "ms..."
    Else
        sbStatus.SimpleText = "Playing (" & sNote & ") " & fFrequency & "hz for " & fDuration & "ms..."
    End If
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Save"
            mnuFileSave_Click
        Case "Remove"
            mnuFileRemove_Click
        Case "Properties"
            mnuFileProperties_Click
        Case "Convert"
            mnuFileConvert_Click
        Case "Export"
            mnuFileExport_Click
        Case "Options"
            mnuFileOptions_Click
        Case "Play"
            mnuPlayPlay_Click
        Case "Stop"
            mnuPlayStop_Click
        Case "Help"
            mnuHelpAbout_Click
        Case Else
            MsgBox "Toolbar function " & Button.Key & " not implemented!", vbExclamation
    End Select
End Sub

Private Sub txtDestination_Change()
    mnuFileExport.Enabled = (txtDestination.Text <> sEmpty)
    tbToolbar.Buttons("Export").Enabled = mnuFileExport.Enabled
End Sub
