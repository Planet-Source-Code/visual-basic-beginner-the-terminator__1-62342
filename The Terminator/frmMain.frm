VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "The Terminator"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "&Terminate All Checked Processes"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   3615
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   4920
      Width           =   3855
   End
   Begin VB.TextBox txtStatus 
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3480
      Width           =   7455
   End
   Begin MSComctlLib.ListView lvwProcesses 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCheckAll 
         Caption         =   "&Check All"
      End
      Begin VB.Menu mnuUncheckAll 
         Caption         =   "&Uncheck All"
      End
      Begin VB.Menu mnuTerminate 
         Caption         =   "&Terminate All Checked Items"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort in Alphabetical Order"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RightMouseClick As Boolean

'Check if any new processes have been added or removed
Private Sub cmdRefresh_Click()
lvwProcesses.ListItems.Clear

Select Case GetTheVersion()

Case 1 'Windows 95/98

Dim f As Long, sname As String
Dim hSnap As Long, proc As PROCESSENTRY32
hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If hSnap = hNull Then Exit Sub
proc.dwSize = Len(proc)
' Iterate through the processes
f = Process32First(hSnap, proc)
Do While f
sname = StrZToStr(proc.szExeFile)

sname = Replace(sname, "\??\", "")
sname = Replace(sname, "\SystemRoot\", "C:\Windows\")
lvwProcesses.ListItems.Add(, , sname).SubItems(1) = proc.th32ProcessID

f = Process32Next(hSnap, proc)
Loop

Case 2 'Windows NT

Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim i As Long
'Get the array containing the process id's for each process object
cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
 cb = cb * 2
 ReDim ProcessIDs(cb / 4) As Long
 lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop
NumElements = cbNeeded / 4

For i = 1 To NumElements
'Get a handle to the Process
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
'Got a Process handle
If hProcess <> 0 Then
'Get an array of the module handles for the specified
'process
lRet = EnumProcessModules(hProcess, Modules(1), 200, _
cbNeeded2)
'If the Module Array is retrieved, Get the ModuleFileName
If lRet <> 0 Then
ModuleName = Space(MAX_PATH)
nSize = 500
lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
ModuleName = Replace(ModuleName, "\??\", "")
ModuleName = Replace(ModuleName, "\SystemRoot\", "C:\WINDOWS\")
If ProcessIDs(i) <> GetCurrentProcessId Then
lvwProcesses.ListItems.Add(, , Left(ModuleName, lRet)).SubItems(1) = ProcessIDs(i)
End If
End If
End If
'Close the handle to the process
lRet = CloseHandle(hProcess)
Next

End Select
End Sub

'Goto mnuTerminate_Click
Private Sub cmdTerminate_Click()
mnuTerminate_Click
End Sub

Private Sub Form_Load()
'Add the columns
lvwProcesses.ColumnHeaders.Add , , "Filename"
lvwProcesses.ColumnHeaders.Add , , "Process ID"
cmdRefresh_Click
End Sub

'If the form is resized, resize the controls
Private Sub Form_Resize()
On Error GoTo ERROR_HANDLER
lvwProcesses.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdRefresh.Height - txtStatus.Height
LV_AutoSizeColumn lvwProcesses
lvwProcesses.ColumnHeaders.Item(2).Width = Me.ScaleWidth - lvwProcesses.ColumnHeaders.Item(1).Width
txtStatus.Move 0, lvwProcesses.Height, Me.ScaleWidth, txtStatus.Height
cmdTerminate.Move 0, lvwProcesses.Height + txtStatus.Height, Me.ScaleWidth / 2, cmdRefresh.Height
cmdRefresh.Move Me.ScaleWidth / 2, lvwProcesses.Height + txtStatus.Height, Me.ScaleWidth / 2, cmdRefresh.Height
ERROR_HANDLER:
End Sub

Private Sub lvwProcesses_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader.Index = 1 Then lvwProcesses.Sorted = True
End Sub

'If an item is clicked, if the item is unchecked then it will be unchecked and vice versa
Private Sub lvwProcesses_ItemClick(ByVal Item As MSComctlLib.ListItem)
If RightMouseClick = False Then
If Item.Checked = False Then
Item.Checked = True
Else
Item.Checked = False
End If
End If
RightMouseClick = False
End Sub

'For not checking or unchecking the item clicked if lvwProcesses is right clicked
Private Sub lvwProcesses_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then RightMouseClick = True
End Sub

'If the user right clicks lvwProcesses then show the popup mnuFile
Private Sub lvwProcesses_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuFile
End Sub

'Show frmAbout
Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

'Check all items in lvwProcesses
Private Sub mnuCheckAll_Click()
Dim i As Long
For i = 1 To lvwProcesses.ListItems.Count
lvwProcesses.ListItems.Item(i).Checked = True
Next
End Sub

Private Sub mnuSort_Click()
lvwProcesses.SortKey = 1
lvwProcesses.Sorted = True
End Sub

Private Sub mnuTerminate_Click()
Dim i As Long
Dim NumChecked As Long
Dim ProcessId As Long

'Loop through all of the items to make that that at least one checkbox is checked
For i = 1 To lvwProcesses.ListItems.Count
If lvwProcesses.ListItems.Item(i).Checked = True Then NumChecked = NumChecked + 1
Next
If NumChecked = 0 Then MsgBox "Please check one or more checkboxes next to the items that you want to terminate", vbCritical, ""

'Loop through all items and try to terminate all of the checked items
For i = 1 To lvwProcesses.ListItems.Count

If lvwProcesses.ListItems.Item(i).Checked = True Then

'Try to terminate the process using OpenProcess and TerminateProcess
ProcessId = OpenProcess(PROCESS_ALL_ACCESS, False, Val(lvwProcesses.ListItems(i).SubItems(1)))
If TerminateProcess(ProcessId, 0) = 0 Then
txtStatus.Text = txtStatus.Text & "Method 1(OpenProcess+TerminateProcess) failed to terminate " & lvwProcesses.ListItems(i).Text & vbCrLf
CloseHandle ProcessId
Else
txtStatus.Text = txtStatus.Text & "Method 1(OpenProcess+TerminateProcess) successfully terminated " & lvwProcesses.ListItems(i).Text & vbCrLf
CloseHandle ProcessId
End If

'Try to terminate the process using DebugActiveProcess
If DebugActiveProcess(Val(lvwProcesses.ListItems(i).SubItems(1))) = 0 Then
txtStatus.Text = txtStatus.Text & "Method 2(DebugActiveProcess) failed to terminate " & lvwProcesses.ListItems(i).Text & vbCrLf
Else
txtStatus.Text = txtStatus.Text & "Method 2(DebugActiveProcess) successfully terminated " & lvwProcesses.ListItems(i).Text & ". Please exit this program to completely terminate the process." & vbCrLf
End If

'Try to terminate the process using a function that I found online
If KillProcess(Val(lvwProcesses.ListItems(i).SubItems(1)), 0) = False Then
txtStatus.Text = txtStatus.Text & "Method 3(Custom) failed to terminate " & lvwProcesses.ListItems(i).Text & vbCrLf
Else
txtStatus.Text = txtStatus.Text & "Method 3(Custom) successfully terminated " & lvwProcesses.ListItems(i).Text & ". Please exit this program to completely terminate the process." & vbCrLf
End If

txtStatus.Text = txtStatus.Text & vbCrLf

End If

Next

End Sub

'Uncheck all items in lvwProcesses
Private Sub mnuUncheckAll_Click()
Dim i As Long
For i = 1 To lvwProcesses.ListItems.Count
lvwProcesses.ListItems.Item(i).Checked = False
Next
End Sub
