VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMSVolume 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Levels"
   ClientHeight    =   3105
   ClientLeft      =   2610
   ClientTop       =   2070
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   1980
   Begin ComctlLib.Slider sliderMasterVolume 
      Height          =   1395
      Left            =   210
      TabIndex        =   2
      Top             =   570
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   2461
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   630
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2505
      Width           =   675
   End
   Begin VB.TextBox txtWaveOutVolume 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   255
      Width           =   555
   End
   Begin VB.TextBox txtMasterVolume 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   255
      Width           =   555
   End
   Begin ComctlLib.Slider sliderWaveOutVolume 
      Height          =   1395
      Left            =   1155
      TabIndex        =   3
      Top             =   570
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   2461
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin VB.Label lblWaveOutVolume 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Wave Out"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1950
      Width           =   765
   End
   Begin VB.Label lblMasterVolume 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Master"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   1950
      Width           =   495
   End
End
Attribute VB_Name = "frmMSVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      
Dim hmixer As Long          ' mixer handle
Dim VolCtrl As MIXERCONTROL ' master volume control
Dim WavCtrl As MIXERCONTROL ' wave output volume control
Dim rc As Long              ' return code
Dim ok As Boolean           ' boolean return code

Private Sub Form_Load()
    Dim Volume As Long
    
    'Open the mixer with deviceID 0.
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer."
        Exit Sub
    End If
    ' Get the Master (Speaker) volume control
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)
    If (ok = True) Then
        ' If the function successfully gets the volume control,
        ' then get the current setting of the control.
        Volume = GetVolumeControlValue(hmixer, VolCtrl)
        'Then use current setting to position the slider
        'display it in the text box.
        If Volume <> -1 Then
            txtMasterVolume.Text = Volume
            sliderMasterVolume.Value = 65535 - Volume
        End If
    End If
    ' Get the Wave Output volume control
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
    'And do the same to the WaveOut controls.
    If (ok = True) Then
        Volume = GetVolumeControlValue(hmixer, WavCtrl)
        If Volume <> -1 Then
            txtWaveOutVolume.Text = Volume
            sliderWaveOutVolume.Value = 65535 - Volume
        End If
    End If
End Sub
      
Private Sub cmdExit_Click()
    'Close the mixer with mixer handle.
    rc = mixerClose(hmixer)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't close the mixer."
    End If
    Unload Me
    End
End Sub

Private Sub sliderMasterVolume_Scroll()
    Dim Volume As Long
    Volume = 65535 - CLng(sliderMasterVolume.Value)
    txtMasterVolume.Text = Volume
    SetVolumeControl hmixer, VolCtrl, Volume
End Sub

Private Sub sliderWaveOutVolume_Scroll()
    Dim Volume As Long
    Volume = 65535 - CLng(sliderWaveOutVolume.Value)
    txtWaveOutVolume.Text = Volume
    SetVolumeControl hmixer, WavCtrl, Volume
End Sub
