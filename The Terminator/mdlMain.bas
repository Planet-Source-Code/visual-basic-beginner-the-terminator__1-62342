Attribute VB_Name = "mdlMain"
Option Explicit

'For ListView AutoSize
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LVM_FIRST = &H1000

Public Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long

Public Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long

Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long ' This process
th32DefaultHeapID As Long
th32ModuleID As Long ' Associated exe
cntThreads As Long
th32ParentProcessID As Long ' This process's parent process
pcPriClassBase As Long ' Base priority of process threads
dwFlags As Long
szExeFile As String * 260 ' MAX_PATH
End Type

Public Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long '1 = Windows 95.
  '2 = Windows NT

szCSDVersion As String * 128
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle _
    As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias _
    "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
    ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal _
    TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
    PreviousState As Any, ReturnLength As Any) As Long

'This isn't my function. I found it somewhere online
Function KillProcess(ByVal hProcessID As Long, Optional ByVal ExitCode As Long) As Boolean
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
    
    ' Windows NT/2000 require a special treatment
    ' to ensure that the calling process has the
    ' privileges to shut down the system
    
    ' under NT the high-order bit (that is, the sign bit)
    ' of the value retured by GetVersion is cleared
    If GetVersion() >= 0 Then
        ' open the tokens for the current process
        ' exit if any error
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY, hToken) = 0 Then
            GoTo CleanUp
        End If
        
        ' retrieves the locally unique identifier (LUID) used
        ' to locally represent the specified privilege name
        ' (first argument = "" means the local system)
        ' Exit if any error
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            GoTo CleanUp
        End If
    
        ' complete the TOKEN_PRIVILEGES structure with the # of
        ' privileges and the desired attribute
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
    
        ' try to acquire debug privilege for this process
        ' exit if error
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, _
            ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
    End If
    
    ' now we can finally open the other process
    ' while having complete access on its attributes
    ' exit if any error
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        ' call was successful, so we can kill the application
        ' set return value for this function
        KillProcess = (TerminateProcess(hProcess, ExitCode) <> 0)
        ' close the process handle
        CloseHandle hProcess
    End If
    
    If GetVersion() >= 0 Then
        ' under NT restore original privileges
        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
        
CleanUp:
        If hToken Then CloseHandle hToken
    End If
End Function










Function StrZToStr(s As String) As String
StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function GetTheVersion() As Long
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
GetTheVersion = osinfo.dwPlatformId
End Function

'Automatically resize the columns so that no text gets cut off
Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
Dim C As ColumnHeader
If Column Is Nothing Then
For Each C In LV.ColumnHeaders
SendMessage LV.hwnd, LVM_FIRST + 30, C.Index - 1, -1
Next
Else
SendMessage LV.hwnd, LVM_FIRST + 30, Column.Index - 1, -1
End If
LV.Refresh
End Sub
