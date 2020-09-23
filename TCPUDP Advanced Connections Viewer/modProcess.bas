Attribute VB_Name = "modProcess"
Option Explicit

Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetLastError Lib "kernel32.dll" () As Long
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long

      Public Type PROCESSENTRY32
         dwSize As Long
         cntUsage As Long
         th32ProcessID As Long           ' This process
         th32DefaultHeapID As Long
         th32ModuleID As Long            ' Associated exe
         cntThreads As Long
         th32ParentProcessID As Long     ' This process's parent process
         pcPriClassBase As Long          ' Base priority of process threads
         dwFlags As Long
         szExeFile As String * 260       ' MAX_PATH
      End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long           '1 = Windows 95.
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
Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Public Const LANG_NEUTRAL As Long = &H0

Type Process1
    ProcessName As String
    pID As Long
End Type

Public Procs() As Process1

Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function getVersion() As Long
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

getVersion = osinfo.dwPlatformId
End Function

Sub GetProcesses()

Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim NumElements2 As Long
Dim lRet As Long
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
        
    ReDim Procs(NumElements)
        
    For i = 1 To NumElements
        Procs(i).pID = ProcessIDs(i)
        Procs(i).ProcessName = ReturnProcessExe(ProcessIDs(i))
    Next
End Sub

Public Function ReturnProcessExe(pID As Long) As String
Dim hProcess As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim cbNeeded2 As Long
Dim nSize As Long
Dim ModuleName As String

'Get a handle to the Process
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pID)
'Got a Process handle
If hProcess <> 0 Then
    'Get an array of the module handles for the specified process
    lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
    'If the Module Array is retrieved, Get the ModuleFileName
    If lRet <> 0 Then
        ModuleName = Space(MAX_PATH)
        nSize = 500
        lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
        ReturnProcessExe = Left(ModuleName, lRet)
    End If
    lRet = CloseHandle(hProcess)
End If

End Function

Public Function TerminateProcessById(lProcessID As Long) As Boolean
  Dim hProcess As Long
  Dim lngReturn As Long
    
    hProcess = OpenProcess(1&, -1&, lProcessID)
    lngReturn = TerminateProcess(hProcess, 0&)
    
    If lngReturn = 0 Then
        RetrieveError
        Exit Function
    End If
    
    TerminateProcessById = True
End Function

'Public Function ProcessID2hWnd(pID As Long) As Long
'Dim hWnd As Long
'    Call GetWindowThreadProcessId(hWnd, pID)
'    ProcessID2hWnd = hWnd
'End Function

Public Function hWnd2ProcessID(hWnd As Long) As Long
Dim pID As Long
    Call GetWindowThreadProcessId(hWnd, pID)
    hWnd2ProcessID = pID
End Function

Private Sub RetrieveError()
  Dim strBuffer As String
    
    'Create a string buffer
    strBuffer = Space(200)
    
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strBuffer, 200, ByVal 0&
    'Show the message
    MsgBox strBuffer, vbApplicationModal + vbExclamation + vbApplicationModal, "Error!"
End Sub

'Sub SuspendThreadbyID(lProcessID As Long, bSuspend As Boolean)
'Select Case bSuspend
'    Case True
'        SuspendThread (ProcessID2hWnd(lProcessID))
'    Case False
'        ResumeThread (ProcessID2hWnd(lProcessID))
'End Select
'End Sub

