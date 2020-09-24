VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "B2SP (W32.Sircam.Worm@mm) Virus Remover"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About Program"
      Height          =   405
      Left            =   3480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1665
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete Virus"
      Height          =   405
      Left            =   1560
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1665
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Height          =   405
      Left            =   5400
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1305
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.DirListBox dirList 
      Height          =   1890
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.FileListBox filList 
      Height          =   1845
      Left            =   6240
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00404040&
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2760
      Width           =   7575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   6960
   End
   Begin VB.Image imgDeleteRegKey 
      Height          =   240
      Left            =   4320
      Picture         =   "Form1.frx":0442
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgChangeDefaultURL 
      Height          =   240
      Left            =   4320
      Picture         =   "Form1.frx":0544
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   4200
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgDeleteRegEntry 
      Height          =   240
      Left            =   480
      Picture         =   "Form1.frx":0646
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgClearRegValue 
      Height          =   240
      Left            =   480
      Picture         =   "Form1.frx":0748
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSearchVBS 
      Height          =   240
      Left            =   480
      Picture         =   "Form1.frx":084A
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   360
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblInProgress 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Press Delete Virus To Remove..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2190
      TabIndex        =   11
      Top             =   210
      Width           =   4035
   End
   Begin VB.Label lblClearRegValue 
      BackColor       =   &H00404040&
      Caption         =   "Edit registry keys..."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   2865
   End
   Begin VB.Label lblChangeDefaultURL 
      BackColor       =   &H00404040&
      Caption         =   "Edit Internet Explorer startup page..."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4800
      TabIndex        =   9
      Top             =   1560
      Width           =   2865
   End
   Begin VB.Label lblSearchVBS 
      BackColor       =   &H00404040&
      Caption         =   "Search and delete virus..."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   2865
   End
   Begin VB.Label lblDeleteRegEntry 
      BackColor       =   &H00404040&
      Caption         =   "Remove added registry keys..."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   2865
   End
   Begin VB.Label lblDeleteRegKey 
      BackColor       =   &H00404040&
      Caption         =   "Edit and remove damaged registry files..."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   1080
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'B2SP Security Softwares
'Decleration

Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SecurityAttributes
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal mKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal mKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    
    
    ' --------------------------------------------------------------------------------
    ' Re RegSetValueEx: If you declare the lpData parameter as String, you must
    ' pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    ' --------------------------------------------------------------------------------
    
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal szValuename As String, ByVal lpReserved As Long, _
    ByVal dwValuetype As Long, bData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal szValuename As String, ByVal lpReserved As Long, _
    ByVal dwValuetype As Long, dwData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" _
    (ByVal mKey As Long, ByVal szValuename As String, ByVal lpReserved As Long, _
    ByVal dwValuetype As Long, ByVal szData As String, ByVal cbData As Long) As Long
    
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal mKey As Long, ByVal lpSubKey As String) As Long
    
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal mKey As Long, ByVal lpValueName As String) As Long
    
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal mKey As Long) As Long

    ' =================================================================================
    ' (The following are for use during code testing only, e.g. to create fictitious
    ' entries and inspect them, and clear them at the end).

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
    (ByVal mKey As Long, ByVal szSubkey As String, ByVal lpReserved As Long, _
    ByVal szClass As String, ByVal dwOptions As Long, ByVal dwDesiredAccess As Long, _
    lpSecurityAttributes As SecurityAttributes, lphResult As Long, _
    lpdwDisposition As Long) As Long
    
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
    lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal mKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long
    
Private Const OPTION_NON_VOLATILE = &H0    ' Info is stored in a file and is preserved
    ' =================================================================================


Private Const FILE_ATTRIBUTE_NORMAL = &H80
 
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

' Reg key security attribute
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_CREATE_SUBKEY = &H4&
Private Const KEY_ENUMERATE_SUBKEY = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20
Private Const READ_CONTROL = &H20000
Private Const WRITE_OWNER = &H80000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
      KEY_ENUMERATE_SUBKEY Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUBKEY

Private Const REG_NONE = 0&
Private Const REG_SZ = 1&                ' Unicode null terminated string
Private Const REG_BINARY = 3             ' Binary
Private Const REG_DWORD = 4              ' 32-bit number
Private Const REG_DWORD_BIG_ENDIAN = 5

Dim arrFileNames(5)                      ' 5 elements array
Dim arrRegClearValue(1, 1)               ' 2 elements and 2 dimension array
Dim arrRegDeleteEntry(1, 1)
Dim arrRegDeleteKey(3, 1)
Dim mCurrUserSubKey As String
Dim mUserSubKey As String
Dim mStartPageEntry As String
Dim mStartPageValue As String
Dim mCurrUserSubKey2 As String
Dim mStartPageEntry2 As String
Dim mStartPageValue2 As String
Dim mSearchPattern As String
Dim mStopFlag As Boolean
Dim mRegHandle As Long
Dim mCount As Integer
Dim mresult
Dim OldFilePathName As String
Dim NewFilePathName As String
Dim regedit
Private WSHShell As Object
Private Sub Command1_Click()
MsgBox "Zain.A Tawfiq @ B2SP Security @ http://security.b2sp.net - b2sp@b2sp.net"
End Sub

Private Sub Form_Load()
      ' File to delete
    arrFileNames(0) = "sircam.sys"
    arrFileNames(1) = "W32.Sircam"
    arrFileNames(2) = "Worm@mm worm"
    arrFileNames(3) = "Worm@mm"

      ' The following are all under HKEY_LOCAL_MACHINE (a mRegHandle here)
      ' Values to clear
    arrRegClearValue(0, 0) = "\Software\Microsoft\Windows\CurrentVersion\RunServices"
    arrRegClearValue(0, 1) = "Driver32"
      ' Entries to delete
    arrRegDeleteEntry(0, 0) = "\Software\Microsoft\Windows\CurrentVersion\RunServices"
    arrRegDeleteEntry(0, 1) = "Driver32"
    
      ' Key to delete
    arrRegDeleteKey(0, 0) = "\Software\Microsoft\Windows\CurrentVersion\RunServices"
    arrRegDeleteKey(0, 1) = "Driver32"
    arrRegDeleteKey(2, 0) = "\Software\Microsoft\Windows\CurrentVersion\RunServices"
    arrRegDeleteKey(2, 1) = "Driver32"
    
    
      ' Default URL to
    mCurrUserSubKey = "Software\Microsoft\Internet Explorer\Main"
    mUserSubKey = "Software\Microsoft\Internet Explorer\Main"
    mStartPageEntry = "Start Page"
    mStartPageValue = "http://security.b2sp.net"
    
    'edit registry value
    mCurrUserSubKey2 = "exefile\shell\open\command"
    mStartPageEntry2 = "(Default)"
    mStartPageValue2 = "%" & "1" & "%" & "*"
    
    
      ' File to search for...
    mSearchPattern = "sircam.sys"
End Sub
'for deleting registry keys
Sub regdelete(regkey)
Set regedit = CreateObject("WScript.Shell")
regedit.regdelete regkey
End Sub
' for editing registry keys
Sub regcreate(regkey, regvalue)
Set regedit = CreateObject("WScript.Shell")
regedit.RegWrite regkey, regvalue
End Sub
Private Sub cmdProceed_Click()
lblInProgress.Caption = "Searching For Virus Files..."
    On Error Resume Next 'to resume erros
    
    regdelete "HKEY_LOCAL_MACHINE\Software\SirCam" 'delete the registry hey that the virus created (all OS)
    Call Kill("C:\Recycled\Sirc32.exe") 'delete virus files...if any errors that will not occure (all OS)
    Call Kill("%System%\Scam32.exe") 'delete virus files...if any errors that will not occure (all OS)
    Call Kill("C:\winnt\Scam32.exe") 'delete virus files...if any errors that will not occure (win2000)
    Call Kill("C:\Windows\Scam32.exe") 'delete virus files...if any errors that will not occure (win9x & 95 & 98 & ME)
    Call Kill("C:\Win98\Scam32.exe") 'delete virus files...if any errors that will not occure (win98)
    Call Kill("C:\Windows\Temp\Scam32.exe") 'delete virus files...if any errors that will not occure (win9x & 95 & 98 & ME)
    Call Kill("C:\Winnt\Temp\Scam32.exe") 'delete virus files...if any errors that will not occure (win2000)
    OldFilePathName = "C:\Winnt\system32\run32.exe" 'rename the system file for win2k that the virus renamed
NewFilePathName = "C:\Winnt\system32\rundll32.exe" 'rename it to the original name... win2k

Name OldFilePathName As NewFilePathName 'rename

    OldFilePathName = "C:\Windows\system32\run32.exe" 'the same as the one above but for (win9x & 95 & 98 & ME)
NewFilePathName = "C:\Windows\system32\run32.exe" 'same as the above but for (win9x & 95 & 98 & ME)

Name OldFilePathName As NewFilePathName 'rename

    cmdProceed.Enabled = False
    
      ' Delete registry key
    DoDeleteRegKey
      '----------------------------------------------------------
    
      ' Clear and report the values of registry entries if found
    DoNullRegEntryValue
      ' Delete and report the registry entries if found
    DoDeleteRegEntry
      '----------------------------------------------------------
    'change registry values
    DoChangecommand
      ' Change and report the default start page
    DoChangeDefaultURL
      ' Search whole disk for files with filespec of "*.???.VBS"
    mStopFlag = False
    DoSearchVBS
    lblInProgress.Caption = "Virus Removed!"
End Sub
'search for virus file and get path
Private Sub RunFilesInDir(inDir)
    On Error Resume Next
    Dim mfile As String
    Dim mFileName As String
    Dim i As Integer
    For i = 0 To UBound(arrFileNames)
         mfile = arrFileNames(i)
         mFileName = Dir$(inDir & "\" & mfile)
         If mFileName <> "" Then
              SetFileAttributes mFileName, FILE_ATTRIBUTE_NORMAL
              Kill inDir & "\" & mfile
              Text2.Text = Text2.Text & "  " & mFileName & vbCrLf
              mCount = mCount + 1
         End If
    Next i
End Sub
' for all registry edition and delete and add.... this will do every thing
Private Sub DoNullRegEntryValue()
    On Error GoTo errHandler
    Dim mStr As String
    Dim i As Integer
    
    Text2.Text = Text2.Text & vbCrLf & "Registry files cleaned:" & vbCrLf
    mRegHandle = HKEY_LOCAL_MACHINE
    mCount = 0
    For i = 0 To UBound(arrRegClearValue)
        mStr = GetRegEntry(mRegHandle, arrRegClearValue(i, 0), arrRegClearValue(i, 1))
        If mStr <> "" Then
            mCount = mCount + 1
            Text2.Text = Text2.Text & "  HKEY_LOCAL_MACHINE\" & arrRegClearValue(i, 0) & "\" & _
               arrRegClearValue(i, 1) & vbCrLf
            SetRegEntry mRegHandle, arrRegClearValue(i, 0), arrRegClearValue(i, 1), ""
        End If
    Next i
    If mCount = 0 Then
        Text2.Text = Text2.Text & "Nothing Found!" & vbCrLf
    End If
    
    imgClearRegValue.Visible = True
    Exit Sub
errHandler:
    ErrMsgProc "DoNullRegEntryValue"
End Sub
'delete the registry entries
Private Sub DoDeleteRegEntry()
    Dim mSubkey As String
    Dim mEntry As String
    Dim mStr As String
    Dim i As Integer
    
    Text2.Text = Text2.Text & vbCrLf & "Deleted registry files:" & vbCrLf
    mRegHandle = HKEY_LOCAL_MACHINE
    mCount = 0
    For i = 0 To UBound(arrRegDeleteEntry)
        mStr = GetRegEntry(mRegHandle, arrRegDeleteEntry(i, 0), arrRegDeleteEntry(i, 1))
        If mStr <> "" Then
             mCount = mCount + 1
             Text2.Text = Text2.Text & "  HKEY_LOCAL_MACHINE\" & arrRegDeleteEntry(i, 0) & _
                "\" & arrRegDeleteEntry(i, 1) & vbCrLf
             DelRegEntry mRegHandle, arrRegDeleteEntry(i, 0), arrRegDeleteEntry(i, 1)
        End If
    Next i
    If mCount = 0 Then
         Text2.Text = Text2.Text & "Nothing Found!" & vbCrLf
    End If
    
    imgDeleteRegEntry.Visible = True
    Exit Sub
errHandler:
    ErrMsgProc "DoDeleteRegEntry"
End Sub
    
    
    
' Delete Registry keys
Private Sub DoDeleteRegKey()
    Dim mKey As Long
    Dim mSub As String
    Dim One_Level_Up As String
    Dim mSubsub As String
    Dim i As Integer
    
    Text2.Text = Text2.Text & "Deleted from registry:" & vbCrLf
    mRegHandle = HKEY_LOCAL_MACHINE
    mCount = 0
    For i = 0 To UBound(arrRegDeleteKey)
        mSub = arrRegDeleteKey(i, 0) & "\" & arrRegDeleteKey(i, 1)
        mresult = RegOpenKeyEx(mRegHandle, mSub, 0, KEY_ALL_ACCESS, mKey)
        If mresult = 0 Then
            One_Level_Up = arrRegDeleteKey(i, 0)
            mSubsub = arrRegDeleteKey(i, 1)
            mresult = RegOpenKeyEx(mRegHandle, One_Level_Up, 0, KEY_ALL_ACCESS, mKey)
            If mresult = 0 Then
                 mCount = mCount + 1
                 Text2.Text = Text2.Text & "  HKEY_LOCAL_MACHINE\" & mSub & vbCrLf
                 RegDeleteKey mKey, mSubsub
                 RegCloseKey mKey
            End If
        End If
    Next i
    If mCount = 0 Then
         Text2.Text = Text2.Text & "Nothing Found!" & vbCrLf
    End If
    
    imgDeleteRegKey.Visible = True
    Exit Sub
errHandler:
    ErrMsgProc "DoDeleteRegkey"
End Sub
'to order the registry entry so the program can edit or delte
Private Function GetRegEntry(ByVal inMainKey As Long, ByVal inSubKey As String, ByVal inEntry As String) As String
    On Error Resume Next
    Dim mKey As Long
    Dim mBuffer As String * 255
    Dim mBufSize As Long
    mresult = RegOpenKeyEx(inMainKey, inSubKey, 0, KEY_READ, mKey)
    If mresult = 0 Then
          mBufSize = Len(mBuffer)
          mresult = RegQueryValueEx(mKey, inEntry, 0, REG_SZ, mBuffer, mBufSize)
          If mresult = 0 Then
                If mBuffer <> "" Then
                     GetRegEntry = Mid$(mBuffer, 1, mBufSize)
                End If
                RegCloseKey mKeys
          Else
                GetRegEntry = ""
          End If
    Else
          GetRegEntry = ""
    End If
End Function
'set the registry entries
Private Sub SetRegEntry(ByVal inMainKey As Long, ByVal inSubKey As String, ByVal inEntry As String, ByVal inValue As String)
    On Error Resume Next
    Dim mKey As Long
    mresult = RegOpenKeyEx(inMainKey, inSubKey, 0, KEY_WRITE, mKey)
    If mresult = 0 Then
         mresult = RegSetValueExString(mKey, inEntry, 0, REG_SZ, inValue, Len(inValue))
         RegCloseKey mKey
    End If
End Sub
'delete the reigstry entry that the virus created
Private Sub DelRegEntry(ByVal inMainKey As Long, ByVal inSubKey As String, ByVal inEntry As String)
    On Error Resume Next
    Dim mKey As Long
    mresult = RegOpenKeyEx(inMainKey, inSubKey, 0, KEY_ALL_ACCESS, mKey)
    If mresult = 0 Then
           ' NB key must be closed for proper deletion
         RegCloseKey mKey
         mresult = RegDeleteValue(mKey, inEntry)
'         RegCloseKey mKey
    End If
End Sub
'change the startopage for internet explorer
Private Sub DoChangeDefaultURL()
    On Error GoTo errHandler
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & vbCrLf & "Internet Explorer startpage changed to:" & vbCrLf
    Text2.Text = Text2.Text & "  " & mStartPageValue
    mRegHandle = HKEY_CURRENT_USER
    SetRegEntry mRegHandle, mCurrUserSubKey, mStartPageEntry, mStartPageValue
    mRegHandle = HKEY_USERS
    SetRegEntry mRegHandle, mUserSubKey, mStartPageEntry, mStartPageValue
    imgChangeDefaultURL.Visible = True
    Exit Sub
errHandler:
    ErrMsgProc "ChangeDefaultURL"
End Sub
'adding new key because the virus deleted the onld one...
Private Sub DoChangecommand()
    On Error GoTo errHandler
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & vbCrLf & "Internet Explorer startpage changed to:" & vbCrLf
    Text2.Text = Text2.Text & "  " & mStartPageValue
    mRegHandle = HKEY_CLASSES_ROOT
    SetRegEntry mRegHandle, mCurrUserSubKey2, mStartPageEntry2, mStartPageValue2
    Exit Sub
errHandler:
    ErrMsgProc "ChangeDefaultURL"
End Sub
'this is just the simple search @ the end
Private Sub DoSearchVBS()
    On Error GoTo errHandler
    Dim tmp
    Text2.Text = Text2.Text & vbCrLf
    Text2.Text = Text2.Text & vbCrLf & "Virus Files:" & vbCrLf
      ' Ensure check whole drive
    dirList.Path = "\"
    DoEvents
    mCount = 0
    SearchIt
    If mCount = 0 Then
         Text2.Text = Text2.Text & "No Virus Files!" & vbCrLf
    End If
    imgSearchVBS.Visible = True
    Exit Sub
errHandler:
    ErrMsgProc "SearchVBS"
End Sub
'this is also commands for searching
Private Sub SearchIt()
    Dim mFirstPath As String
    Dim mErrDirDiver As Boolean
    Dim mDirCount As Integer
    Dim mNumFiles As Integer
    
     ' Perform recursive search.
     ' Update dirList.Path if it is different from the currently
     ' selected directory, otherwise perform the search.
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub
    End If

       ' Continue with the search.
    filList.Pattern = mSearchPattern

    mFirstPath = dirList.Path
    mDirCount = dirList.ListCount

      ' Start recursive direcory search.
    mNumFiles = 0              ' Reset found files indicator
    mErrDirDiver = DirDiver(mFirstPath, mDirCount, "")
    
      ' Recursive direcory search ended
      
     ' If user clicks Stop meanwhile, don't continue,
    If mStopFlag = True Then
        Exit Sub
    End If
    
    If mErrDirDiver = True Then
        Exit Sub
    End If
    filList.Path = dirList.Path
End Sub
'also for searching the files for the virus
Private Function DirDiver(NewPath As String, mDirCount As Integer, BackUp As String) As Integer
     ' If user clicks to stop, then stop
    If mStopFlag Then
        Exit Function
    End If

      ' Recursively search directories from NewPath down...
      ' NewPath is searched on this recursion.
      ' BackUp is origin of this recursion.
      ' mDirCount is number of subdirectories in this directory.
    Dim mDirToPeek As Integer
    Dim mAbandon As Integer
    Dim mOldPath As String
    Dim mCurrPath As String
    Dim mEntry As String
    Dim i As Integer
    
    DirDiver = False             ' Assumed first. Set to False if there is an error.
    
    DoEvents
    If mStopFlag Then
        DirDiver = True
        Exit Function
    End If
    
    On Local Error GoTo errHandler:
    
    mDirToPeek = dirList.ListCount    ' How many directories below this?
    
    Do While mDirToPeek > 0 And mStopFlag = False
        mOldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(mDirToPeek - 1)
            mAbandon = DirDiver((dirList.Path), mDirCount%, mOldPath)
        End If
        ' Go up one level in directories.
        mDirToPeek = mDirToPeek - 1
        If mAbandon = True Then
            mStopFlag = True
            Exit Function
        End If
    Loop
    
      ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then                ' Check for 2 bytes/character
            mCurrPath = dirList.Path                  ' If at root level, leave as is...
        Else
            mCurrPath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For i = 0 To filList.ListCount - 1            ' Show conforming files.
            mEntry = mCurrPath + filList.List(i)
            Text2.Text = Text2.Text & "  " & mEntry & vbCrLf
            Call Kill(mEntry)
            mCount = mCount + 1
       Next i
    End If
    If BackUp <> "" Then                              ' If there is a superior dir, move it.
        dirList.Path = BackUp
    End If
    Exit Function
errHandler:
End Function
'thats very simple for drive changing (drive list box)
Private Sub DrvList_Change()
    On Error GoTo errHandler
    dirList.Path = drvList.Drive
    Exit Sub
errHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub
'when changing the directory list box
Private Sub DirList_Change()
    filList.Path = dirList.Path
    filList.Pattern = mSearchPattern
End Sub
'same as above
Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub
'close program
Private Sub cmdExit_Click()
    mStopFlag = True
    DoEvents

Unload Me
End
End Sub
'the error handler
Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub



'===================================================================================
' END OF CODE
' please vote for me :)
'===================================================================================

