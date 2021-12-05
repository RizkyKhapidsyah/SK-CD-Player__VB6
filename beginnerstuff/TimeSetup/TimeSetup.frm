VERSION 5.00
Begin VB.Form TimeSetup 
   Caption         =   "Setup Time"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "TimeSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "TimeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim lpSystemTime As SYSTEMTIME
    Dim lpLocalTime As SYSTEMTIME
    Dim TZ As TIME_ZONE_INFORMATION
    Dim X As Integer

Private Sub Form_Load()
    '*****************************************************************************
    'First get the System Time.  Since it is supposed to be UTC/Zulu/GMT time
    'and is calculated by the system based on the Local Time plus any Time Zone
    'offset, we need to get it before we go mucking about with the Time Zone
    'Information.  And since we will be setting the Time Zone to GMT, we need
    'to set the Local Time to GMT before we change the Time Zone.
    
    GetSystemTime lpSystemTime              'Get what's supposedly GMT
  
    lpLocalTime.wYear = lpSystemTime.wYear  'Transfer to Local Time structure
    lpLocalTime.wMonth = lpSystemTime.wMonth
    lpLocalTime.wDayOfWeek = lpSystemTime.wDayOfWeek
    lpLocalTime.wDay = lpSystemTime.wDay
    lpLocalTime.wHour = lpSystemTime.wHour
    lpLocalTime.wMinute = lpSystemTime.wMinute
    lpLocalTime.wSecond = lpSystemTime.wSecond
    lpLocalTime.wMilliseconds = lpSystemTime.wMilliseconds
    
    SetLocalTime lpLocalTime            'And set the Local Time to what
                                        'should be UTC.  That is, provided
                                        'the Time Zone and Local Time were
                                        'set correctly to begin with.
    '*****************************************************************************
    'Now we set the Time Zone of the system to UTC/GMT/Zulu.  Note
    'that the actual Locale where Greenwich, England is located uses
    'Daylight Savings Time but that the other "Zero Hour" time zone
    'called "Monrovia,Casablanca" is also GMT but without Daylight
    'Savings Time.  So we'll use that one.  Strangely the system lists
    'it in the system Registry under
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones
    'as Greenwich while listing the actual Greenwich, England location as GMT.
    'Ah, the minds? of the Microsoft Programmers.
    
    TZ.Bias = 0                         'Time Zone Bias
    Erase TZ.StandardName               'StandardName
    For X = 1 To Len("Greenwich Standard Time")
        TZ.StandardName(X - 1) = Asc(Mid("Greenwich Standard Time", X, 1))
    Next X
    TZ.StandardDate.wYear = 0           'StandardDate
    TZ.StandardDate.wMonth = 0
    TZ.StandardDate.wDayOfWeek = 0
    TZ.StandardDate.wDay = 0
    TZ.StandardDate.wHour = 0
    TZ.StandardDate.wMinute = 0
    TZ.StandardDate.wSecond = 0
    TZ.StandardDate.wMilliseconds = 0
    TZ.StandardBias = 0                 'StandardBias (None)
    
    Erase TZ.DaylightName               'DaylightName
    For X = 1 To Len("Greenwich Daylight Time")
        TZ.DaylightName(X - 1) = Asc(Mid("Greenwich Daylight Time", X, 1))
    Next X
    TZ.DaylightDate.wYear = 0           'DaylightDate
    TZ.DaylightDate.wMonth = 0
    TZ.DaylightDate.wDayOfWeek = 0
    TZ.DaylightDate.wDay = 0
    TZ.DaylightDate.wHour = 0
    TZ.DaylightDate.wMinute = 0
    TZ.DaylightDate.wSecond = 0
    TZ.DaylightDate.wMilliseconds = 0
    TZ.DaylightBias = 0                 'DaylightBias (None)
       
    SetTimeZoneInformation TZ           'Set it
    
    'Now to change the Time Format to the 24 hour mode.
    
    'First we gotta change the WIN.INI File 'cause the stupid clock part
    'of TimeDate.cpl uses the WIN.INI to get it's settings. (16 bit stuff?)
    WriteProfileString "intl", "iTime", "1"     '24 hour format
    WriteProfileString "intl", "iTLZero", "1"   'Leading zero
    WriteProfileString "intl", "s1159", ""      'No AM
    WriteProfileString "intl", "s2359", ""      'No PM
    WriteProfileString "intl", "sTime", ":"     'Time separator
    
    'Change the Regional Settings (Locale)
    SetLocaleInfo GetSystemDefaultLCID(), LOCALE_STIMEFORMAT, "HH:mm:ss"
    
    'Tell the system about the change
    SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
    
    Unload Me
    End
End Sub




