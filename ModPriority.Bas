Option Explicit
' Get and Set File Priority - Edgemeal Feb.2006
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION As Long = &H200
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function OpenProcess Lib "KERNEL32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32.dll" (ByVal Handle As Long) As Long

Public Enum Priorities
    RealTime = 256       'REALTIME_PRIORITY
    High = 128           '.HIGH_PRIORITY
    Above_Normal = 32768 '.ABOVE_NORMAL
    Normal = 32          '.NORMAL_PRIORITY
    Below_Normal = 16384 '.BELOW_NORMAL
    Low = 64             '.IDLE_PRIORITY
End Enum

Public Function GetPriority(PID As Long) As Long
    Dim hwnd As Long
    hwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, PID)
    If hwnd Then
        GetPriority = GetPriorityClass(hwnd)
        CloseHandle hwnd
    End If
End Function

Public Sub SetPriority(PID As Long, PriCls As Priorities)
    Dim priHwnd As Long
    'If GetPriority(PID) = PriCls Then Exit Sub 'If Priority already set then just exit.
    priHwnd = OpenProcess(PROCESS_SET_INFORMATION, False, PID)
    If priHwnd Then
        SetPriorityClass priHwnd, PriCls
        CloseHandle priHwnd
    End If
End Sub

Public Function PriorityToString(Pri As Long) As String
    Dim S As String
    Select Case Pri
        Case 256 'REALTIME_PRIORITY
            S = " Real Time"
        Case 128 '.HIGH_PRIORITY
            S = "High"
        Case 32768  '.ABOVE_NORMAL
            S = "Above Normal"
        Case 32 '.NORMAL_PRIORITY
            S = "Normal"
        Case 16384 '.BELOW_NORMAL
            S = "Below Normal"
        Case 64  '.IDLE_PRIORITY
            S = "Low"
    End Select
    PriorityToString = S
End Function

Public Function StringToPriority(Pri As Priorities) As Long
    StringToPriority = Pri
End Function