VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NXMisc Demo"
   ClientHeight    =   6615
   ClientLeft      =   1155
   ClientTop       =   1545
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNXMiscDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBrowse 
      Caption         =   "BrowseForFolder"
      Height          =   1095
      Left            =   5520
      TabIndex        =   56
      Top             =   4800
      Width           =   2775
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtBrowse 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblBrowse 
         Alignment       =   2  'Center
         Caption         =   "Click here to select a folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   405
         MouseIcon       =   "frmNXMiscDemo.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   720
         Width           =   1965
      End
   End
   Begin VB.Frame fraUsername 
      Caption         =   "Computer Username"
      Height          =   615
      Left            =   5520
      TabIndex        =   53
      Top             =   4080
      Width           =   2775
      Begin VB.CommandButton cmdUsername 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblUsername 
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fraUptime 
      Caption         =   "Computer Uptime"
      Height          =   735
      Left            =   5520
      TabIndex        =   50
      Top             =   3240
      Width           =   2775
      Begin VB.Timer tmrUptime 
         Interval        =   512
         Left            =   2280
         Top             =   240
      End
      Begin VB.CommandButton s 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblUptime 
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fraComputerName 
      Caption         =   "Computer Name"
      Height          =   615
      Left            =   5520
      TabIndex        =   47
      Top             =   2520
      Width           =   2775
      Begin VB.CommandButton cmdComputerName 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblComputerName 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fraToShortFilename 
      Caption         =   "ToShortFilename"
      Height          =   1095
      Left            =   5520
      TabIndex        =   43
      Top             =   1320
      Width           =   2775
      Begin VB.TextBox txtToShortFilenameOut 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "C:\PROGRA~1"
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdToShortFilename 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtToShortFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Text            =   "C:\Program Files"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraToLongFilename 
      Caption         =   "ToLongFilename"
      Height          =   1095
      Left            =   5520
      TabIndex        =   39
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtToLongFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Text            =   "C:\PROGRA~1"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdToLongFilename 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtToLongFilenameOut 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "C:\Program Files"
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame fraCleanFilename 
      Caption         =   "CleanFilename"
      Height          =   1095
      Left            =   2640
      TabIndex        =   35
      Top             =   5400
      Width           =   2775
      Begin VB.TextBox txtCleanFilenameOut 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "C:\autoexec.bat"
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdCleanFilename 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtCleanFilenameIn 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Text            =   "file:///C:/autoexec.bat"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraUnhex 
      Caption         =   "Unhex"
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   5400
      Width           =   2415
      Begin VB.TextBox txtUnhex 
         Height          =   285
         Left            =   120
         MaxLength       =   7
         TabIndex        =   33
         Text            =   "Enter a hex value here."
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdUnhex 
         Caption         =   "?"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblUnhex 
         Caption         =   "Numerical Form: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame fraTextStrip 
      Caption         =   "TextStrip"
      Height          =   2295
      Left            =   2640
      TabIndex        =   22
      Top             =   3000
      Width           =   2775
      Begin VB.TextBox txtTextStripOut 
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtTextStripIn 
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtTextStripAllowed 
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Text            =   "aeiouy"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdTextStrip 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblTextStripOut 
         Caption         =   "Stripped Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblTextStripIn 
         Caption         =   "Source Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblTextStripAllowed 
         Caption         =   "Allowed Characters:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraShutdown 
      Caption         =   "Shut Down"
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   2415
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "?"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdDoShutdown 
         Caption         =   "Do It"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optShutdown 
         Caption         =   "Shut down the computer"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optShutdown 
         Caption         =   "Restart the computer"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optShutdown 
         Caption         =   "Log off user"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Frame fraWinVer 
      Caption         =   "Windows Version"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
      Begin VB.CommandButton cmdWinVer 
         Caption         =   "?"
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblWinVer 
         Caption         =   "0.0.0000"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraScreenRes 
      Caption         =   "Screen Resolution"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   2415
      Begin VB.CommandButton cmdScreenRes 
         Caption         =   "?"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblScreenRes 
         Caption         =   "0x0"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraGetDrive 
      Caption         =   "GetDriveInfo/Letter"
      Height          =   2775
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdGetDrive 
         Caption         =   "?"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtGetDrive 
         BackColor       =   &H8000000F&
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ListBox lstGetDrive 
         Height          =   840
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraCrash 
      Caption         =   "Crash"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
      Begin VB.CommandButton cmdDoCrash 
         Caption         =   "Crash This Program"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdCrash 
         Caption         =   "?"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame fraBint 
      Caption         =   "Bint"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtBint 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Enter a value here."
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdBint 
         Caption         =   "?"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblBint 
         Caption         =   "BINT Result: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Caption         =   "modNXMisc.bas and NXMisc Demo (C) 2001-2002 by Mark Christian."
      Height          =   375
      Left            =   5520
      TabIndex        =   60
      Top             =   6000
      Width           =   2775
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
MsgBox "modNXMisc.bas and NXMisc Demo were programmed by Mark Christian in 2001. The source code is freely distributable." & vbNewLine & vbNewLine & "Contact: mark.christian@bigfoot.com", 64
End Sub

Private Sub cmdBint_Click()
MsgBox "The BINT() function works in the same manner as the INT() function, and can in fact take its place in any action." & vbNewLine & "Its power comes from its ability to handle strings and convert them to integers. BINT strips all non-numerical characters  and any extra decimal points from the string before converting them to Long, rounding appropriately, and returning the result as an Integer." & vbNewLine & vbNewLine & "Usage:" & vbNewLine & "BINT (number)", 64, "Bint - Better INT"
End Sub



Private Sub cmdBrowse_Click()
MsgBox "The BrowseForFolder function brings up a standard Windows folder selection dialog box and returns the selected path as a string. It accepts two parameters: the owner form's hWnd, and the prompt. If the owner form's hWnd is present, the dialog is show modally on top of the owner. If not, the dialog is shown normally. The prompt appears at the top of the dialog box, and is set to " & Chr(34) & "Please select a folder." & Chr(34) & " if it is omitted." & vbNewLine & vbNewLine & "Usage:" & vbNewLine & "Folder = BrowseForFolder (SomeForm.hWnd, Prompt)", vbInformation, "BrowseForFolder"
End Sub

Private Sub cmdCleanFilename_Click()
MsgBox "CleanFilename parses the filename you pass to it and cleans it up, converting from Internet-style file:/// paths as necessary, removing double-slashes, converting front-slashes to back-slashes, and getting rid of redundant strings (such as \.\).", 64, "CleanFilename"
End Sub

Private Sub cmdComputerName_Click()
MsgBox "ComputerName is a function that is designed to be accessed like a variable. It simply returns a string containing the name of the computer." & vbNewLine & vbNewLine & "Usage:" & vbNewLine & "ComputerName" & vbNewLine & vbNewLine & "Example:" & vbNewLine & "MsgBox " & Chr(34) & "Your computer is named " & Chr(34) & " & ComputerName & " & Chr(34) & "." & Chr(34) & ", vbInformation", vbInformation, "ComputerName"
End Sub

Private Sub cmdCrash_Click()
MsgBox "The Crash function crashes the application or development environment." & vbNewLine & "Its use has yet to be determined, but it is interesting.", 64, "Crash"
End Sub

Private Sub cmdDoCrash_Click()
Crash
End Sub

Private Sub cmdDoFocusForm_Click()
MsgBox "This fun"
End Sub


Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGetDrive_Click()
MsgBox "The GetDriveList and GetDriveInfo commands are used in tandem here." & vbNewLine & "The list of drive letters is obtained from GetDriveList, and the specific drive information is used by passing the drive letter to GetDriveInfo.", 64, "GetDriveList and GetDriveInfo"
End Sub

Private Sub cmdScreenRes_Click()
MsgBox "The GetScreenSize function returns the screen resolution. Depending on the argument sent to it, it can return the height, width, or both." & vbNewLine & vbNewLine & "Examples:" & vbNewLine & "GetScreenSize 0 - " & GetScreenSize(0) & vbNewLine & "GetScreenSize 1 - " & GetScreenSize(1) & vbNewLine & "GetScreenSize 2 - " & GetScreenSize(2), 64, "GetScreenRes"
End Sub

Private Sub cmdDoShutdown_Click()
If optShutdown(0).Value = True Then LogOff
If optShutdown(1).Value = True Then Restart
If optShutdown(2).Value = True Then Shutdown
End Sub

Private Sub cmdShutdown_Click()
MsgBox "These actions are carried out by the Logoff, Restart, and Shutdown commands. They require no arguments. The Username function is also employed here to set the text 'Log off (username)'." & vbNewLine & "Note that all three of these commands simply pass the proper arguments to the DoShutdown function automatically. See the comments in DoShutdown for more details.", 64, "DoShutdown"
End Sub


Private Sub cmdTextStrip_Click()
MsgBox "The TextStrip command provides an easy way to modify user input. It goes through the string passed to it and removes any characters not specified in the optional Allowed argument. By default, the allowed characters are only lowercase letters, uppercase letters, and numbers." & vbNewLine & "It should be noted that the Bint function uses a modified version of the TextStrip code, and many of the other included functions make use of the actual TextStrip function.", 64
End Sub

Private Sub cmdToLongFilename_Click()
MsgBox "ToLongFilename accepts a short filename as a string argument and returns the full Windows filename." & vbNewLine & vbNewLine & "If the file does not exist, the returned string will be empty.", 64, "ToLongFilename"
End Sub

Private Sub cmdToShortFilename_Click()
MsgBox "ToShortFilename accepts a full Windows filename filename as a string argument and returns the shortened equivalent.." & vbNewLine & vbNewLine & "If the file does not exist, the returned string will be empty.", 64, "ToShortFilename"
End Sub


Private Sub cmdUnhex_Click()
MsgBox "Unhex provides an easy command to convert Base-16 (hexadecimal) numbers stored in strings into their Base-10 equivalents. Its usage is similar to the conversion functions built into Visual Basic (such as String(), Int(), CLng(), etc)." & vbNewLine & "Note that due to the limits of the Long data type, the Unhex command will return -1 for numbers longer than seven characters (268,435,455 in Base-10).", 64, "Unhex"

End Sub

Private Sub cmdUptime_Click()
MsgBox "Uptime returns a string containing uptime information, that is, how long the computer has been running. If you omit the includeComputerName parameter or set it to TRUE, the function will return a string that looks like this: " & vbNewLine & vbNewLine & Uptime & vbNewLine & vbNewLine & "If you set includeComputerName to FALSE, it will look like this:" & vbNewLine & vbNewLine & Uptime(False), vbInformation, "Uptime"
End Sub

Private Sub cmdUsername_Click()
MsgBox "Username is a function that is designed to be accessed like a variable. It simply returns a string containing the name of the currently logged in user." & vbNewLine & vbNewLine & "Usage:" & vbNewLine & "Username" & vbNewLine & vbNewLine & "Example:" & vbNewLine & "MsgBox " & Chr(34) & "Your username is " & Chr(34) & " & Username & " & Chr(34) & "." & Chr(34) & ", vbInformation", vbInformation, "ComputerName"
End Sub

Private Sub cmdWinVer_Click()
MsgBox "The GetWindowsVersion function returns the current Windows version number." & vbNewLine & "You are running Windows v" & GetWindowsVersion & ".", 64, "GetWindowsVersion"
End Sub

Private Sub Form_Load()
'This code is used in the GetDrive demonstration
  lstGetDrive.Clear
  For i = 0 To Len(GetDriveLetter) - 1
    lstGetDrive.AddItem GetDriveLetter(i) & ":"
  Next i
  lstGetDrive.ListIndex = 0

lblScreenRes.Caption = GetScreenSize
lblWinVer.Caption = GetWindowsVersion
optShutdown(0).Caption = "Log off " & Username
txtTextStripIn.Text = "Let's show vowels only!"
lblComputerName.Caption = ComputerName
lblUptime.Caption = Uptime(False)
lblUsername.Caption = Username
End Sub


Private Sub lblBrowse_Click()
txtBrowse.Text = BrowseForFolder(Me.hwnd)
If txtBrowse.Text = "" Then 'the user cancelled
  txtBrowse.Text = "User pressed cancel."
End If
End Sub

Private Sub lblBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBrowse.ForeColor = &H80000011 'Disabled text color constant
End Sub


Private Sub lblBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBrowse.ForeColor = vbHighlight
End Sub


Private Sub lstGetDrive_Click()
'This sends the selected drive letter to the GetDriveInfo function and puts the result into txtGetDrive.Text.
'Note that the information displayed can be customized to a specific field, and custom reports can be
'generated from the distinct parts. See the notes in the GetDriveInfo function.
txtGetDrive.Text = GetDriveInfo(Left(lstGetDrive.Text, 1))
End Sub


Private Sub tmrUptime_Timer()
lblUptime.Caption = Uptime(False)
End Sub

Private Sub txtBint_Change()
X = Bint(txtBint.Text)
If (lblBint.Caption <> "BINT Result: " & X) Then lblBint.Caption = "BINT Result: " & X
End Sub


Private Sub txtCleanFilenameIn_Change()
txtCleanFilenameOut.Text = CleanFilename(txtCleanFilenameIn.Text)
End Sub

Private Sub txtTextStripAllowed_Change()
txtTextStripOut.Text = TextStrip(txtTextStripIn.Text, txtTextStripAllowed.Text)
End Sub

Private Sub txtTextStripIn_Change()
txtTextStripOut.Text = TextStrip(txtTextStripIn.Text, txtTextStripAllowed.Text)
End Sub


Private Sub txtToLongFilename_Change()
On Error Resume Next
txtToLongFilenameOut.Text = ToLongFilename(txtToLongFilename.Text)
End Sub

Private Sub txtToShortFilename_Change()
On Error Resume Next
txtToShortFilenameOut.Text = ToShortFilename(txtToShortFilename.Text)
End Sub


Private Sub txtUnhex_Change()
On Error Resume Next
X = Unhex(txtUnhex.Text)
If (lblUnhex.Caption <> "Numerical Form: " & X) Then lblUnhex.Caption = "Numerical Form: " & X
End Sub


