VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Memory Stats"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   255
      Left            =   8280
      TabIndex        =   37
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame5 
      Caption         =   "Computer name"
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   2760
      Width           =   4215
      Begin VB.Label lblCompName 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Lock Workstation"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Username"
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   4215
      Begin VB.Label lblUser 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "System Info"
      Height          =   1695
      Left            =   4440
      TabIndex        =   20
      Top             =   2760
      Width           =   4215
      Begin VB.Label lblSysInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label lblSysInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblSysInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblSysInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label lblSysInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox picPGBar 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   15
         Top             =   480
         Width           =   3735
         Begin VB.Shape shpStatus 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            DrawMode        =   5  'Not Copy Pen
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   4
            Left            =   0
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Drive Space"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   3735
         End
         Begin VB.Label lblInfo2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0gb Of 0gb"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   17
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.ComboBox cboDrive 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblSerialOut 
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblSer 
         Caption         =   "Serial:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblFileSysOut 
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblFS 
         Caption         =   "File System:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblVolOut 
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblVol 
         Caption         =   "Volume:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblTypeOut 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblType 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblDrive 
         Caption         =   "Drive:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Memory Stats"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox picPGBar 
         AutoRedraw      =   -1  'True
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   1
         Top             =   360
         Width           =   3735
         Begin VB.Shape shpStatus 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            DrawMode        =   5  'Not Copy Pen
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblInfo2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0kb Of 0kb"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Virtual Memory"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.PictureBox picPGBar 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   240
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   10
         Top             =   280
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Shape shpStatus 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            DrawMode        =   5  'Not Copy Pen
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   3
            Left            =   0
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Memory Load"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.Timer timStats 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3600
         Top             =   1320
      End
      Begin VB.PictureBox picPGBar 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   7
         Top             =   1080
         Width           =   3735
         Begin VB.Shape shpStatus 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            DrawMode        =   5  'Not Copy Pen
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblInfo2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0kb Of 0kb"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Physical Memory"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.PictureBox picPGBar 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   4
         Top             =   1800
         Width           =   3735
         Begin VB.Shape shpStatus 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            DrawMode        =   5  'Not Copy Pen
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Page File"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   3735
         End
         Begin VB.Label lblInfo2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0kb Of 0kb"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   3735
         End
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type DISKSPACE
  dsRoot As String
  dsFreeBytesAvailToCaller As Currency
  dsTotalBytes As Currency
  dsTotalFreeBytes As Currency
  dsDriveType As String
  dsVolume As String
  dsSerial As Long
  dsFileSystem As String
End Type

Private Type MEMORYSTATUS
  dwLength As Long
  dwMemoryLoad As Long
  dwTotalPhys As Long
  dwAvailPhys As Long
  dwTotalPageFile As Long
  dwAvailPageFile As Long
  dwTotalVirtual As Long
  dwAvailVirtual As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Dim PId As String
Dim OSInfo As OSVERSIONINFO

Private Sub cboDrive_Click()
  GetSpace
End Sub

Private Sub cmdAbout_Click()
  ShellAbout Me.hWnd, App.Title, "Copyright (C) 2001 Camalot Designs" & vbCrLf & "Visit: http://camalot.virtualave.net", ByVal 0&
End Sub

Private Sub cmdLock_Click()
  'LockWorkStation
  End
End Sub

Private Sub Form_Load()
  Dim dwLen As Long
  Dim strString As String
  Dim strSave As String, ret As Long, x As Integer
  Dim strUserName As String
  
    
  strSave = String(255, Chr$(0))
  ret& = GetLogicalDriveStrings(255, strSave)
  For x = 1 To 100
    If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
    cboDrive.AddItem Left$(strSave, InStr(1, strSave, Chr$(0)) - 1)
    strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
  Next x
  cboDrive.ListIndex = 1
  
  'Create a buffer
  dwLen = MAX_COMPUTERNAME_LENGTH + 1
  strString = String(dwLen, "X")
  'Get the computer name
  GetComputerName strString, dwLen
  'get only the actual data
  strString = Left(strString, dwLen)
  lblCompName.Caption = strString
  
  
  strUserName = String(100, Chr$(0))
  'Get the username
  GetUserName strUserName, 100
  'strip the rest of the buffer
  strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
  lblUser.Caption = strUserName
  
  GetStats
  GetSysInfo
  
  'If OSInfo.dwPlatformId <> 2 Then
    'cmdLock.Enabled = False
  'End If
  cmdLock.Caption = "E&xit System Stats"
  timStats.Enabled = True
End Sub
Private Function Percent(Total As Long, stepnum As Long, output As Long) As Long
  On Error Resume Next
  stepnum& = stepnum& / 100
  Total& = Total& / 100
  Percent = CLng(stepnum& * output& / Total&)
  stepnum& = stepnum& * 100
  Total& = Total& * 100
End Function

Private Sub GetSpace()
  Dim dSpace As DISKSPACE
  Dim ret As Long
  Dim dsPer As Long
  Dim dsAvail As Long
  Dim dsFree As Long
  Dim dsTotal As Long
  'Dim Serial As Long, VName As String, FSName As String
  
  dSpace.dsRoot = cboDrive.List(cboDrive.ListIndex) 'Store the Drive
  lblTypeOut.Caption = "Checking..."
  DoEvents 'Allow the caption to change
  
  'Create buffers
  dSpace.dsVolume = String$(255, Chr$(0))
  dSpace.dsFileSystem = String$(255, Chr$(0))
  'Get the volume information
  GetVolumeInformation dSpace.dsRoot, dSpace.dsVolume, 255, dSpace.dsSerial, 0, 0, dSpace.dsFileSystem, 255
  'Strip the extra chr$(0)'s
  dSpace.dsVolume = Left$(dSpace.dsVolume, InStr(1, dSpace.dsVolume, Chr$(0)) - 1)
  dSpace.dsFileSystem = Left$(dSpace.dsFileSystem, InStr(1, dSpace.dsFileSystem, Chr$(0)) - 1)
 
  lblVolOut.Caption = dSpace.dsVolume
  lblFileSysOut.Caption = dSpace.dsFileSystem
  lblSerialOut.Caption = dSpace.dsSerial
   
  'Get The Space Info and store it in dSpace
  ret = GetDiskFreeSpaceEx(dSpace.dsRoot, dSpace.dsFreeBytesAvailToCaller, dSpace.dsTotalBytes, dSpace.dsTotalFreeBytes)
  
  'Get the drive type
  ret = GetDriveType(dSpace.dsRoot)
  Select Case ret
    Case 2
        dSpace.dsDriveType = "Removable Drive"
    Case 3
        dSpace.dsDriveType = "Fixed Drive"
    Case 4
        dSpace.dsDriveType = "Remote Drive"
    Case 5
        dSpace.dsDriveType = "CD-Rom Drive"
    Case 6
        dSpace.dsDriveType = "Ram Disk"
    Case Else
        dSpace.dsDriveType = "Unrecognized"
  End Select
  'Display the drive type
  lblTypeOut.Caption = dSpace.dsDriveType
  
  dsTotal = dSpace.dsTotalBytes * 1024
  dsFree = dSpace.dsTotalFreeBytes * 1024
  dsAvail = dSpace.dsFreeBytesAvailToCaller
  'Show the percent of FREE Space
  dsPer = Percent(dsTotal, dsFree, lblInfo(4).Width)
  shpStatus(4).Width = dsPer
  
  Debug.Print dSpace.dsRoot, Percent(dsTotal, dsFree, 100) & "% Free"
  
  'Set the Caption for the lblInfo
  lblInfo(4).Caption = "Free Disk Space (" & Percent(dsTotal, dsFree, 100) & "%)"
  
  
  'Determine what type of conversion we should use by the size of the drive
  'Select Case dsTotal
    'Case Is >= 107374182 'A Gig or Bigger Use Giga Byte
      lblInfo2(4).Caption = ConvertGigaBytes(dsFree, 2) & " of " & ConvertGigaBytes(dsTotal, 2)
      Debug.Print ConvertGigaBytes(dsFree, 2), dsFree
    'Case Is < 107374182 And dsTotal > 1048576 'less than a Giga Byte so use mega byte
      'lblInfo2(4).Caption = ConvertMegaBytes(dsAvail, 2) & " of " & ConvertMegaBytes(CLng(dsTotal), 2)
      'Debug.Print ConvertMegaBytes(dsAvail, 2), dsAvail
    'Case Is < 1048576 'less then a megabyte so use kilo byte
      'lblInfo2(4).Caption = ConvertKiloBytes(dsAvail, 2) & " of " & ConvertKiloBytes(CLng(dsTotal), 2)
      'Debug.Print ConvertKiloBytes(dsAvail, 2), dsAvail
    'Case Is < 1024 'less then kilo byte so use byte
      'lblInfo2(4).Caption = ConvertBytes(dsAvail, 2) & " of " & ConvertBytes(CLng(dsTotal), 2)
      'Debug.Print ConvertBytes(dsAvail, 2), dsAvail
  'End Select
  
  
End Sub

Private Sub GetSysInfo()
  Dim SInfo As SYSTEM_INFO
  
  Dim ret As Long
  
  'Get Version Info
  OSInfo.dwOSVersionInfoSize = Len(OSInfo)
  ret& = GetVersionEx(OSInfo)
  
  Select Case OSInfo.dwPlatformId
    Case 0
      PId = "Windows 32s"
    Case 1
      PId = "Windows 95/98/ME"
    Case 2
      PId = "Windows NT/2000/XP"
  End Select

  
  
  'Get the system information
  GetSystemInfo SInfo

  lblSysInfo(0).Caption = "Number Of Proccessors: " & Str$(SInfo.dwNumberOrfProcessors)
  lblSysInfo(1).Caption = "Proccessor Type: " & Str$(SInfo.dwProcessorType)
  lblSysInfo(2).Caption = "OS: " & PId
  lblSysInfo(3).Caption = "Version:" & Str$(OSInfo.dwMajorVersion) & "." + LTrim(Str(OSInfo.dwMinorVersion))
  
  lblSysInfo(4).Caption = "Build: " & Str(OSInfo.dwBuildNumber) & " (" & Left$(OSInfo.szCSDVersion, InStr(1, OSInfo.szCSDVersion, Chr$(0)) - 1) & ")"
  'msgbox osinfo.
End Sub

Private Sub GetStats()
  Dim MemStat As MEMORYSTATUS
    
  Dim lngPhysMem As Long
  Dim lngAPhysMem As Long
  Dim lngPPhysMem As Long
  
  Dim lngVirMem As Long
  Dim lngAVirMem As Long
  Dim lngPVirMem As Long
  
  Dim lngPageFile As Long
  Dim lngAPageFile As Long
  Dim lngPPageFile As Long
  
  'retrieve the memory status
  GlobalMemoryStatus MemStat
  
  lngPhysMem = MemStat.dwTotalPhys '/ 1024
  lngAPhysMem = MemStat.dwAvailPhys '/ 1024
  lngPPhysMem = Percent(lngPhysMem, lngAPhysMem, lblInfo(0).Width)
  lblInfo(0).Caption = "Physical Memory (" & Percent(lngPhysMem, lngAPhysMem, 100) & "%)"
  'Debug.Print lngPPhysMem & "%", "Physical"
  shpStatus(0).Width = lngPPhysMem
  lblInfo2(0).Caption = ConvertMegaBytes(lngAPhysMem, 2) & " Of " & ConvertMegaBytes(lngPhysMem, 2)
  
  
  lngVirMem = MemStat.dwTotalVirtual / 1024
  lngAVirMem = MemStat.dwAvailVirtual / 1024
  lngPVirMem = Percent(lngVirMem, lngAVirMem, lblInfo(1).Width)
  lblInfo(1).Caption = "Virtual Memory (" & Percent(lngVirMem, lngAVirMem, 100) & "%)"
  'Debug.Print lngPVirMem & "%", "Virtual"
  shpStatus(1).Width = lngPVirMem
  
  lblInfo2(1).Caption = ConvertMegaBytes(lngAVirMem, 2) & " Of " & ConvertMegaBytes(lngVirMem, 2)
  lblInfo2(1).Caption = ""
  
  lngPageFile = MemStat.dwTotalPageFile '/ 1024
  lngAPageFile = MemStat.dwAvailPageFile '/ 1024
  lngPPageFile = Percent(lngPageFile, lngAPageFile, lblInfo(2).Width)
  lblInfo(2).Caption = "Page File (" & Percent(lngPageFile, lngAPageFile, 100) & "%)"
  'Debug.Print lngPPageFile & "%", "Pagefile"
  shpStatus(2).Width = lngPPageFile
  lblInfo2(2).Caption = ConvertMegaBytes(lngAPageFile, 2) & " Of " & ConvertMegaBytes(lngPageFile, 2)


  'lblInfo(3).Caption = "Memory Load (" & MemStat.dwMemoryLoad & "%)"
  'shpStatus(3).Width = Percent(100, MemStat.dwMemoryLoad, lblInfo(3).Width)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  timStats.Enabled = False
End Sub


Private Sub timStats_Timer()
  GetStats
End Sub
