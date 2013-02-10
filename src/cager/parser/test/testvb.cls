VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal lngMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
	
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const STILL_ACTIVE = &H103

Private WithEvents m_WaitForm As frmWait
Attribute m_WaitForm.VB_VarHelpID = -1

Private m_ExitCode As Long
Private m_ProgHandle As Long

Public Function Run(ByVal cmd As String, Optional WindowStyle As VBA.VbAppWinStyle = vbNormalFocus, Optional ByVal Wait As Boolean) As Long
    Debug.Print "RUN: " & cmd
    m_ProgHandle = Shell(cmd, WindowStyle)
    
    If m_ProgHandle = 0 Then
        ' Failed
        Run = -1
    Else
        Run = WaitOnProgram()
        Debug.Print "EXIT " & Run
    End If
End Function

Private Function WaitOnProgram() As Long
    Set m_WaitForm = New frmWait
    m_WaitForm.ShowWait 
    
    WaitOnProgram = m_ExitCode
End Function

Private Sub m_WaitForm_Timer()
    Dim iExit As Long
    Dim hProg As Long
    
    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, m_ProgHandle)
    
    GetExitCodeProcess hProg, iExit
    
    CloseHandle hProg
    
    If iExit <> STILL_ACTIVE Then
        m_ExitCode = iExit
        m_WaitForm.Visible = False
        Set m_WaitForm = Nothing
    End If
End Sub
