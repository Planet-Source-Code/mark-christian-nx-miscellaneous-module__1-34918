Attribute VB_Name = "modNXMisc"
'NXMISC.BAS
'Enhanced Functions (C) 2001 by Mark Christian
'Contact: mark.christian@bigfoot.com

'Includes code by:
'David Jarrett (BrowseForFolder)
'Rocky Clark (ToLongFilename/ToShortFilename)

'Note: functions with a ReturnMode argument have constants for all of their
'values declared below. I suggest using them, as they make life a lot easier,
'and ensure that changes in future versions of this module will not 'break'
'any code that calls said functions.

'API Declarations
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpUid As LUID) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)

'Constants: BrowseForFolder (Internal Use Only)
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
'Constants: GetScreenSize (ReturnMode)
Public Const ScreenSize = 0
Public Const ScreenWidth = 1
Public Const ScreenHeight = 2
'Constants: GetDriveInfo (ReturnMode)
Public Const VolumeName = 0
Public Const VolumeSerialNumber = 1
Public Const VolumeType = 2
Public Const VolumeFileSystem = 3
Public Const VolumeCapacity = 4
Public Const VolumeAvailable = 5
'Constants: DoShutdown (Internal Use Only)
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const ERROR_NOT_ALL_ASSIGNED = 1300
Public Const SE_PRIVILEGE_ENABLED = 2
Public Const TOKEN_QUERY = &H8
Public Const TOKEN_ADJUST_PRIVILEGES = &H20

'Types
Public Type BrowseInfo
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Public Type LUID
    lowpart As Long
    highpart As Long
End Type

Public Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Function BrowseForFolder(Optional OwnerForm_hWnd As Long = 0, Optional Prompt As String = "Please select a folder.") As String
'Input: (Optional) hWnd of parent folder, (Optional) Prompt
'Output: Selected folder

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
szTitle = Prompt

tBrowseInfo.hWndOwner = OwnerForm_hWnd
tBrowseInfo.lpszTitle = lstrcat(szTitle, "")
tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End If
BrowseForFolder = sBuffer
End Function


Public Function CleanFilename(Filename As String) As String
'CleanFilename
'Input: Messy filename
'Output: Fixes errors, such as double-slashes, and converts file:/// paths to normal Windows paths

'Convert to lowercase
CleanFilename = LCase(Filename)

'Check for file:///
If Left(Filename, 8) = "file:///" Then
  'Filename is in file:/// format
  CleanFilename = Right(Filename, Len(Filename) - 8)
  CleanFilename = Replace(CleanFilename, "%20", " ")
End If

'Clean-ups to be run on both types
CleanFilename = Replace(CleanFilename, "/", "\")
CycleSlashes:
  CleanFilename = Replace(CleanFilename, "\\", "\")
  If InStr(1, CleanFilename, "\\") > 0 Then GoTo CycleSlashes
  
CycleDots:
  CleanFilename = Replace(CleanFilename, "\.\", "\")
  If InStr(1, CleanFilename, "\\") > 0 Then GoTo CycleDots
  
End Function


Public Function ComputerName() As String
Dim cNameBuffer As String, cNameLength As Long

cNameBuffer = String(255, Chr(0))
cNameLength = 255
Call GetComputerName(cNameBuffer, cNameLength)
ComputerName = Left(cNameBuffer, cNameLength)
End Function

Public Function Uptime(Optional includeComputerName = True) As String
Dim Weeks, Days, Hours, Minutes, Seconds

Seconds = Round(GetTickCount / 1000, 0)
Weeks = (Seconds - (Seconds Mod 604800)) / 604800 '604800 is the number of seconds in a week
Seconds = Seconds - (Weeks * 604800)
Days = (Seconds - (Seconds Mod 86400)) / 86400 '86400 is the number of seconds in a day
Seconds = Seconds - (Days * 86400)
Hours = (Seconds - (Seconds Mod 3600)) / 3600 '3600 is the number of seconds in an hour
Seconds = Seconds - (Hours * 3600)
Minutes = (Seconds - (Seconds Mod 60)) / 60 '60 is the number of seconds in an hour
Seconds = Seconds - (Minutes * 60)

If includeComputerName Then
  Uptime = ComputerName & " has been up for " & Weeks & " weeks, " & Days & " days, " & Hours & " hours, " & Minutes & " minutes, " & Seconds & " seconds."
Else
  Uptime = Weeks & " weeks, " & Days & " days, " & Hours & " hours, " & Minutes & " minutes, " & Seconds & " seconds."
End If
End Function
Public Sub DoShutdown(SType As Integer)
'DoShutdown
'Input: SType -- 1 to shutdown, 2 to restart, 0 to log off
'Output: Performs specified action
'Note: Use the Shutdown, Restart, and LogOff functions instead
    
    Dim tLuid          As LUID
    Dim tTokenPriv     As TOKEN_PRIVILEGES
    Dim tPrevTokenPriv As TOKEN_PRIVILEGES
    Dim lResult        As Long
    Dim lToken         As Long
    Dim lLenBuffer     As Long
    Dim lMode As Long
    
    Select Case SType
        Case 1
            ' Shutdown the computer
            lMode = 1
        Case 2
            ' Reboot the computer
            lMode = 2
        Case 3
            ' Log off and select a different user
            lMode = 0
    End Select
    
    If Not bWindowsNT Then
        Call ExitWindowsEx(lMode, 0)
    Else
        '
        ' Get the access token of the current process.  Get it
        ' with the privileges of querying the access token and
        ' adjusting its privileges.
        '
        lResult = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lToken)
        If lResult = 0 Then
            Exit Sub 'Failed
        End If
        '
        ' Get the locally unique identifier (LUID) which
        ' represents the shutdown privilege.
        '
        lResult = LookupPrivilegeValue(0&, "SeShutdownPrivilege", tLuid)
        If lResult = 0 Then Exit Sub 'Failed
        '
        ' Populate the new TOKEN_PRIVILEGES values with the LUID
        ' and allow your current process to shutdown the computer.
        '
        With tTokenPriv
            .PrivilegeCount = 1
            .Privileges.Attributes = SE_PRIVILEGE_ENABLED
            .Privileges.pLuid = tLuid
        lResult = AdjustTokenPrivileges(lToken, False, tTokenPriv, Len(tPrevTokenPriv), tPrevTokenPriv, lLenBuffer)
        End With
        
        If lResult = 0 Then
            Exit Sub 'Failed
        Else
            If Err.LastDllError = ERROR_NOT_ALL_ASSIGNED Then Exit Sub 'Failed
        End If
        '
        '  Shutdown Windows.
        '
        Call ExitWindowsEx(lMode, 0)
    End If
End Sub
Public Function Bint(Number As Variant) As Integer
'Bint (Better INT)
'Input: Any number (long, double, single, or string)
'Output: Valid, properly rounded integer

On Error Resume Next

validChars = "0123456789."
outBuffer = ""
hasDecimal = False
Number = Str(Number)
For i = 1 To Len(Number)
  x = LCase(Mid(Number, i, 1))
  If InStr(1, validChars, x) > 0 Then
    If x = "." Then
      If hasDecimal = False Then
        outBuffer = outBuffer & x
        hasDecimal = True
      End If
    Else
      outBuffer = outBuffer & x
    End If
  End If
Next i
Bint = Int(Round(CLng(outBuffer), 0))
End Function
Public Function Crash(Optional CrashCode As Long = 0)
'Crash
'Input: Optional error code to crash with
'Output: Crashes the application or VB IDE

FatalExit CrashCode
End Function


Public Function FocusForm(FormName As Form)
'FocusForm
'Input: Form to focus
'Output: Brings form to top of stack

BringWindowToTop FormName.hwnd
End Function

Public Function GetDriveInfo(DriveLetter As String, Optional ReturnMode = -1) As String
'GetDriveInfo
'Inputs: Drive letter, optional ReturnMode determines what to return
'Output: Information on specified drive

Dim sBuffer As String * 255, sFileSystem As String * 25, lSerialNumber As Long, lMaxLength As Long, lFlags As Long, lSectors As Long, lBytes As Long, lFreeClusters As Long, lTotalClusters As Long, dCapacity As Double, dAvailable As Double

'Get volume information
GetVolumeInformation DriveLetter & ":\", sBuffer, Len(sBuffer), lSerialNumber, lMaxLength, lFlags, sFileSystem, Len(sFileSystem)

oVolumeName = UCase(TextStrip(sBuffer))
oSerialNumber = Hex(lSerialNumber)
If Len(oSerialNumber) < 8 Then oSerialNumber = Left("00000000", 8 - Len(oSerialNumber)) & oSerialNumber
oFileSystem = UCase(TextStrip(sFileSystem))

'Get drive type
Select Case GetDriveType(DriveLetter & ":\")
  Case 2
    oDriveType = "Removable"
  Case 3
    oDriveType = "Fixed"
  Case 4
    oDriveType = "Network"
  Case 5
    oDriveType = "CD-ROM"
  Case 6
    oDriveType = "RAM disk"
  Case Else
    oDriveType = "Unknown"
End Select

'Get disk size
GetDiskFreeSpace DriveLetter & ":\", lSectors, lBytes, lFreeClusters, lTotalClusters
dCapacity = lTotalClusters * lSectors
dCapacity = dCapacity * lBytes
dAvailable = lFreeClusters * lSectors
dAvailable = dAvailable * lBytes

'Output information
If ReturnMode = -1 Then GetDriveInfo = "Volume Name: " & oVolumeName & vbNewLine & "Serial Number: " & Left(oSerialNumber, 4) & "-" & Right(oSerialNumber, 4) & vbNewLine & "Drive Type: " & oDriveType & vbNewLine & "File System: " & oFileSystem & vbNewLine & "Capacity: " & dCapacity & " bytes" & vbNewLine & "Available: " & dAvailable & " bytes"
If ReturnMode = 0 Then GetDriveInfo = oVolumeName
If ReturnMode = 1 Then GetDriveInfo = oSerialNumber
If ReturnMode = 2 Then GetDriveInfo = oDriveType
If ReturnMode = 3 Then GetDriveInfo = oFileSystem
If ReturnMode = 4 Then GetDriveInfo = dCapacity
If ReturnMode = 5 Then GetDriveInfo = dAvailable
End Function
Public Function GetDriveLetter(Optional DriveNumber = -1) As String
'GetDriveLetter
'Input: Drive number to return a specific letter, or nothing to return a list
'Output: Drive letter, or list of drives

On Error GoTo ErrHand

Dim sBuffer As String * 200

GetLogicalDriveStrings Len(sBuffer), sBuffer

Drives = UCase(TextStrip(sBuffer, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"))

If DriveNumber = -1 Then
  GetDriveLetter = Drives
Else
  GetDriveLetter = Mid(Drives, (DriveNumber + 1), 1)
End If

Exit Function

ErrHand:
GetDriveLetter = ""
End Function
Public Function GetScreenSize(Optional ReturnMode As Integer) As String
'GetScreenSize
'Inputs: Optional ReturnMode specifies what to return
'Output: Screen resolution

sHeight = Screen.Height / Screen.TwipsPerPixelY
sWidth = Screen.Width / Screen.TwipsPerPixelX
If ReturnMode = 0 Then GetScreenSize = sWidth & "x" & sHeight
If ReturnMode = 1 Then GetScreenSize = sWidth
If ReturnMode = 2 Then GetScreenSize = sHeight
End Function

Public Function Shutdown()
'Shutdown
'Input: None
'Output: Shuts down the computer

DoShutdown 1
End Function

Public Function Restart()
'Restart
'Input: None
'Output: Restarts the computer

DoShutdown 2
End Function


Public Function LogOff()
'LogOff
'Input: None
'Output: Logs out the current user

DoShutdown 0
End Function


Public Function ToShortFilename(LongFilename As String) As String
'Input: Long filename
'Output: Short filename equivalent

Dim InputText As String, OutputText As String * 67
Dim OutputLength As Long

InputText = LongFilename
OutputLength = GetShortPathName(InputText, OutputText, Len(OutputText))
InputText = Left(OutputText, OutputLength)
ToShortFilename = InputText
End Function

Public Function ToLongFilename(ShortFilename As String) As String
'Input: Short filename
'Output: Long filename equivalent

Dim OutputLength As Long
Dim OutputText As String

OutputText = String(1024, " ")
OutputLength = GetLongPathName(ShortFilename, OutputText, Len(OutputText))

If OutputLength > Len(OutputText) Then
    OutputText = String(OutputLength + 1, " ")
    OutputLength = GetLongPathName(ShortFilename, OutputText, Len(OutputText))
End If

If OutputLength > 0 Then
    ToLongFilename = Left(OutputText, OutputLength)
End If
End Function

Public Function Username() As String
'UserName
'Input: None
'Output: Returns name of current user

Dim sBuffer As String * 255

GetUserName sBuffer, 255

Username = Trim(TextStrip(sBuffer))
End Function

Public Function GetWindowsVersion() As String
'GetWindowsVersion
'Input: None
'Output: Return Windows version number

Dim WinInfo As OSVERSIONINFO

WinInfo.dwOSVersionInfoSize = Len(WinInfo)
GetVersionEx WinInfo

GetWindowsVersion = WinInfo.dwMajorVersion & "." & WinInfo.dwMinorVersion & "." & WinInfo.dwBuildNumber
End Function

Public Function TextStrip(Text As String, Optional Allowed As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890") As String
'TextStrip
'Inputs: Text to format, allowable characters
'Output: Text stripped of disallowed characters

outBuffer = ""
For i = 1 To Len(Text)
    x = Mid(Text, i, 1)
    If InStr(1, Allowed & vbNewLine, x) > 0 Then
        outBuffer = outBuffer & x
    End If
Next i

TextStrip = outBuffer
End Function
Public Function Unhex(Number As String) As Long
'Unhex
'Input: Hex number in a string
'Output: Long value of hex number

Unhex = CLng("&H0" & Number)
End Function


