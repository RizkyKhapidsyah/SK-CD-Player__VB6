VERSION 5.00
Begin VB.Form frmMixerGetDevCaps 
   Caption         =   "MixerGetDevCaps"
   ClientHeight    =   4275
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4275
   ScaleWidth      =   5910
   Begin VB.TextBox txtCapsError 
      Height          =   285
      Left            =   3600
      TabIndex        =   21
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox txtDestinationCount 
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtSupportBits 
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtProductName 
      Height          =   285
      Left            =   3600
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtDriverVersion 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtProductID 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtMfgID 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4860
      TabIndex        =   8
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox txtError 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtDeviceID 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtDevices 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Error"
      Height          =   195
      Left            =   2640
      TabIndex        =   22
      Top             =   3120
      Width           =   330
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Destination Count"
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Support Bits"
      Height          =   195
      Left            =   2280
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2280
      TabIndex        =   16
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Driver Version"
      Height          =   195
      Left            =   2280
      TabIndex        =   14
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Product ID"
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Manufacturer's ID"
      Height          =   195
      Left            =   2280
      TabIndex        =   10
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Error"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mixer Handle"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mixer ID"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "# Devices"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMixerGetDevCaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '
    'My first attempt at talking to the Sound card using the MCI.
    'Tom McCandless 5/10/98
    '
Dim uMxId As Long
Dim ReturnCode As Long

Private Sub Form_Load()
    Show
    Call subOpen_mixer
    Call subMixerGetDevCaps
End Sub

Private Sub subOpen_mixer()
    Dim Number_of_Devices As Long
    Dim MixerID As Long
    'Get the number of devices
    Number_of_Devices = mixerGetNumDevs()
    txtDevices.Text = Number_of_Devices
    MixerID = Number_of_Devices - 1
    txtDeviceID.Text = MixerID
    'Open and get the Mixer Handle
    ReturnCode = mixerOpen(uMxId, MixerID, 0, 0, 0)
    Label3.Caption = "Mixer Handle"
    txtData.Text = Hex(uMxId)
    txtError.Text = ReturnCode
End Sub

Private Sub subMixerGetDevCaps()
    Dim pmxcaps As MIXERCAPS
    Dim cbmxcaps As Long

    cbmxcaps = 100

    'Ask for the Mixer Device Caps
    ReturnCode = mixerGetDevCaps(uMxId, pmxcaps, cbmxcaps)
    txtCapsError.Text = ReturnCode
    'Look Up the Manufacturer
    txtMfgID.Text = Manufacturer(pmxcaps.wMid)
    If txtMfgID.Text = "Not Listed" Then txtMfgID.Text = Str(pmxcaps.wMid)
    'Look up the Product
    txtProductID.Text = Product(pmxcaps.wPid)
    If txtProductID.Text = "Not Listed" Then txtProductID.Text = Str(pmxcaps.wPid)
    'Decode the Driver Version
    txtDriverVersion.Text = Val("&h" + Left(Hex(pmxcaps.vDriverVersion), Len(Hex(pmxcaps.vDriverVersion)) - 2)) & "." & Val("&h" + Right(Hex(pmxcaps.vDriverVersion), 2))
    'Misc stuff
    txtProductName.Text = pmxcaps.szPname
    txtSupportBits.Text = pmxcaps.fdwSupport
    txtDestinationCount = pmxcaps.cDestinations
End Sub

Private Sub subClose_mixer()
    'Close the mixer
    ReturnCode = mixerClose(uMxId)
    txtError.Text = ReturnCode
End Sub

Private Sub cmdExit_Click()
    'Exit the application
    Call subClose_mixer
    Unload Me
    End
End Sub

