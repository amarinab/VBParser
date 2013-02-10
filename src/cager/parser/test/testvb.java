package cager.parser.test;

public class testvb{

	private Long OpenProcess ( Long dwAccess, Integer fInherit, Long hObject ) {

/* 
 
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long 
 */
	}

	private Long WaitForSingleObject ( Long hHandle, Long lngMilliseconds ) {

/* 
 
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal lngMilliseconds As Long) As Long 
 */
	}

	private Long CloseHandle ( Long hObject ) {

/* 
 
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long 
 */
	}

	private void Sleep ( Long dwMilliseconds ) {

/* 
 
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 
 */
	}

	private Long GetExitCodeProcess ( Long hProcess, Long lpExitCode ) {

/* 
 	
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long 
 */
	}

	private void Run ( String cmd, VBA.VbAppWinStyle WindowStyle, Boolean Wait ) {

/* 
 
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
 */
	}

	private void WaitOnProgram ( ) {

/* 
 
Private Function WaitOnProgram() As Long
    Set m_WaitForm = New frmWait
    m_WaitForm.ShowWait 
    
    WaitOnProgram = m_ExitCode
End Function 
 */
	}

	private void m_WaitForm_Timer ( ) {

/* 
 
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
 */
	}

	private void Shell ( ) {

/* 
 
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
 */
	}

}
