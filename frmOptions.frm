VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "frmOptions"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVolume 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   17
      Text            =   "15"
      Top             =   2812
      Width           =   375
   End
   Begin VB.TextBox txtRepeat 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   5
      Text            =   "4"
      Top             =   1052
      Width           =   495
   End
   Begin VB.TextBox txtAutoLED 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   14
      Text            =   "2"
      Top             =   2372
      Width           =   495
   End
   Begin VB.TextBox txtAutoBacklight 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   11
      Text            =   "1"
      Top             =   1932
      Width           =   495
   End
   Begin VB.TextBox txtAutoVibration 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   8
      Text            =   "4"
      Top             =   1492
      Width           =   495
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      ItemData        =   "frmOptions.frx":000C
      Left            =   1560
      List            =   "frmOptions.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   620
      Width           =   1155
   End
   Begin VB.CheckBox chkOptimize 
      Alignment       =   1  'Right Justify
      Caption         =   "Optimi&ze Ringtone:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldAutoLED 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   2380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldAutoBacklight 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   1940
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldAutoVibration 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1500
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldRepeat 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1060
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Top             =   2820
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Min             =   1
      Max             =   15
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Volu&me:"
      Height          =   195
      Left            =   960
      TabIndex        =   15
      Top             =   2880
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Repeat:"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1120
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ringtone &Style:"
      Height          =   195
      Left            =   450
      TabIndex        =   1
      Top             =   680
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Auto &Vibration:"
      Height          =   195
      Left            =   495
      TabIndex        =   6
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Auto &Backlight:"
      Height          =   195
      Left            =   450
      TabIndex        =   9
      Top             =   2000
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Auto &LED:"
      Height          =   195
      Left            =   795
      TabIndex        =   12
      Top             =   2440
      Width           =   735
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Options window implementation

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    ApplyOptions
    Hide
End Sub

Private Sub Form_Load()
    cboStyle.ListIndex = 1
    txtAutoBacklight_Change
    txtAutoLED_Change
    txtAutoVibration_Change
    txtRepeat_Change
    txtVolume_Change
End Sub

Private Sub sldAutoBacklight_Scroll()
    txtAutoBacklight.Text = sldAutoBacklight.Value
End Sub

Private Sub sldAutoLED_Scroll()
    txtAutoLED.Text = sldAutoLED.Value
End Sub

Private Sub sldAutoVibration_Scroll()
    txtAutoVibration.Text = sldAutoVibration.Value
End Sub

Private Sub sldRepeat_Scroll()
    txtRepeat.Text = sldRepeat.Value
End Sub

Private Sub sldVolume_Scroll()
    txtVolume.Text = sldVolume.Value
End Sub

Private Sub txtAutoBacklight_Change()
    txtAutoBacklight.Text = Val(txtAutoBacklight.Text)
    sldAutoBacklight.Value = txtAutoBacklight.Text
End Sub

Private Sub txtAutoLED_Change()
    txtAutoLED.Text = Val(txtAutoLED.Text)
    sldAutoLED.Value = txtAutoLED.Text
End Sub

Private Sub txtAutoVibration_Change()
    txtAutoVibration.Text = Val(txtAutoVibration.Text)
    sldAutoVibration.Value = txtAutoVibration.Text
End Sub

Private Sub txtRepeat_Change()
    txtRepeat.Text = Val(txtRepeat.Text)
    sldRepeat.Value = txtRepeat.Text
End Sub

Public Sub ApplyOptions()
    frmMain.rtPlayer.Optimize = True
    frmMain.rtSE.SetOptions cboStyle.ListIndex, sldRepeat.Value, sldAutoVibration.Value, sldAutoBacklight.Value, sldAutoLED.Value, sldVolume.Value
End Sub

Private Sub txtVolume_Change()
    txtVolume.Text = Val(txtVolume.Text)
    sldVolume.Value = txtVolume.Text
End Sub
